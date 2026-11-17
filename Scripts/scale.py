"""
Step 02: scale the prepared case's gens & loads to their targets (from Scripts/area.py or Scripts/topo.py, by mode).

Gens and loads are handled by the same code. Their working tables ("targets") are indexed by ObjectID and hold:
    <field>          current value in the case
    <field>_Target   value to reach (from the GridView hour: area totals, or each element's own value)
    <field>_Seed     value in the prepared case
    <field>_Start    value when scaling started (after closing in, at 0 MW)
    Progress         how far the row has been scaled from start to target, in 1/RESOLUTION units
    HeldReason       '' while the row is scaled. Otherwise, why it is held back for now (it is retried later).
    ExclusionReason  '' while the row is scaled or held. Otherwise, why it was excluded for good.
"""
from dataclasses import dataclass, field
from pathlib import Path
import multiprocessing as mp
import numpy as np
import pandas as pd
from . import pw, prep

# ------------------ Tuning thresholds ------------------
EXCLUDE_VMIN = 0.88             # Hold gens/loads back from scaling on buses below this Vpu...
EXCLUDE_VMAX = 1.12             # ...or above this Vpu.
STATCOM_VPU = 0.85              # Add a STATCOM to the lowest-voltage bus below this Vpu...
STATCOM_VNOM_MIN = 50           # ...with nominal kV above this.
COLLAPSE_VPU = 0.80             # Drop branches between buses above and below this Vpu.
PVQV_DV_LIMIT = 0.10            # Exclude buses whose linear PVQV dV estimate exceeds this.
RESOLUTION = 100                # Smallest scaling step is 1/RESOLUTION of the way from start to target.
INITIAL_STEP = 25               # First step, in 1/RESOLUTION units. Doubles after each step which solves, up to MAX_STEP.
MAX_STEP = 100                  # Halves after each step which diverges, down to 1. A step which still diverges is bisected.
STEP_SOLVES = 2                 # Solves a step must pass to count as converged.
FINAL_MAX_ROUNDS = 20           # Max rounds retrying held gens/loads, once the rest have reached their targets.
# Solve aids which change the network (not to match GridView, only to keep the case solving).
# Shunt switching (prep.adjust_shunts) is always on.
ADD_STATCOMS = True             # Add a STATCOM on the lowest-voltage bus below STATCOM_VPU.
DROP_COLLAPSED = True           # Open branches into sections below COLLAPSE_VPU.
MAX_WORKERS = None              # Parallel PowerWorld instances for gen target tests. None = CPU count.

STATCOM_ID = 'SC'   # ID of STATCOMs added during scaling.

# ------------------ Gen & load definitions ------------------
# Per table: each status field, and the MW/Mvar fields which must be 0 while it is Open.
STATUS = {
    'Gen': {'Status': ['MWSetPoint']},
    'Load': {'Status': ['SMW', 'SMvar', 'IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput'], 'DistStatus': ['DistMWInput', 'DistMvarInput']},
}
# The fields which get scaled.
SCALED = {table: statuses['Status'] for table, statuses in STATUS.items()}
# The (MW, Mvar) fields used to estimate voltage change from PVQV sensitivities.
PVQV_FIELDS = {'Gen': ('MWSetPoint', None), 'Load': ('SMW', 'SMvar')}

def write_fields(table: str) -> list[str]:
    """The fields written back to PowerWorld."""
    return ['BusNum', 'ID', *STATUS[table], *SCALED[table]]

def included(df: pd.DataFrame) -> pd.Series:
    return df['ExclusionReason'] == ''

def exclude(df: pd.DataFrame, mask, reason: str):
    """Excludes rows from scaling. The first recorded reason is kept."""
    df.loc[pd.Series(mask, index=df.index) & included(df), 'ExclusionReason'] = reason

def scaling(df: pd.DataFrame) -> pd.Series:
    return included(df) & (df['HeldReason'] == '')

def held(df: pd.DataFrame) -> pd.Series:
    return included(df) & (df['HeldReason'] != '')

def hold(df: pd.DataFrame, mask, reason: str):
    """Holds rows back from scaling, at their current progress, until they are retried."""
    df.loc[pd.Series(mask, index=df.index) & scaling(df), 'HeldReason'] = reason

# ------------------ Targets ------------------
def test_gen_targets(pw_fp: Path, gen_df: pd.DataFrame) -> pd.DataFrame:
    """
    Tests each generator target individually, and adds a 'Success' column reporting which individual changes solve.
    """
    SimAuto = pw.connect()
    pw.open_case(SimAuto, pw_fp)
    if not pw.solve(SimAuto):
        raise pw.SimAutoError(f'Base case does not solve, cannot test gen targets: {pw_fp}')

    target_df = gen_df.assign(Status=gen_df['Status_Target'], MWSetPoint=gen_df['MWSetPoint_Target'])
    ok_rows = included(target_df).tolist()
    success = []
    for i in range(len(target_df)):
        if not ok_rows[i]:
            success.append(True)    # Excluded rows (the swing among them) are never written.
            continue
        # Set case to target value for this specific element, solve, then roll back.
        with pw.trial(SimAuto, keep_changes=False) as t:
            pw.set_df(SimAuto, 'Gen', target_df.iloc[[i]], write_fields('Gen'))
        success.append(t.ok)

    SimAuto.CloseCase()
    return gen_df.assign(Success=success)

def test_gen_targets_parallel(pw_fp: Path, gen_df: pd.DataFrame) -> pd.DataFrame:
    """
    Taking a set of target MW & Status values for generators, tests to see if each one will solve individually.
    """
    num_workers = max(1, min(MAX_WORKERS or mp.cpu_count(), len(gen_df)))
    df_splits = [gen_df.iloc[i::num_workers] for i in range(num_workers)]

    # Run in parallel:
    with mp.Pool(processes=num_workers) as pool:
        results = pool.starmap(test_gen_targets, [(pw_fp, part) for part in df_splits])

    # Run in series (for debugging):
    # results = [test_gen_targets(pw_fp, part) for part in df_splits]

    return pd.concat(results).sort_values(by='Success')

def report_gen_load_balance(targets: dict[str,pd.DataFrame]) -> pd.DataFrame:
    """
    Prints seed and target gen and load MW in one line, and warns if the target imbalance is high. Only Closed gens and
    loads count (an Open one injects nothing, whatever its MW); dist gen only where its load and itself are Closed.
    """
    gen_df, load_df = targets['Gen'], targets['Load']
    def totals(suffix):
        gen_on, load_on = gen_df['Status' + suffix] == 'Closed', load_df['Status' + suffix] == 'Closed'
        gen_mw = gen_df['MWSetPoint' + suffix].where(gen_on, 0.0).sum()
        dist_mw = load_df['DistMWInput' + suffix].where(load_on & (load_df['DistStatus' + suffix] == 'Closed'), 0.0).sum()
        load_mw = sum(load_df[c + suffix].where(load_on, 0.0).sum() for c in ('SMW', 'IMW', 'ZMW'))
        return [gen_mw, dist_mw, load_mw, gen_mw + dist_mw - load_mw]

    total_delta_df = pd.DataFrame(
        {'Base Case': totals(''), 'Target Case': totals('_Target')}
        ,index=['Gen MW', 'Dist Gen MW', 'Load MW', 'Gen + Dist Gen - Load']
    )
    base, target = total_delta_df['Base Case'], total_delta_df['Target Case']
    print(f'Seed gen {base["Gen MW"]:,.0f} / load {base["Load MW"]:,.0f} MW; '
          f'target gen {target["Gen MW"]:,.0f} / load {target["Load MW"]:,.0f} MW')

    base_total_delta, target_total_delta = total_delta_df.iloc[-1]
    if abs(target_total_delta) > 1.5 * abs(base_total_delta):
        print('WARNING: Gen/Load imbalance in the target case is high. This may cause instability.')

    return total_delta_df

# ------------------ Exclusions ------------------
def compute_pvqv_exclusions(targets, pvqv_df) -> dict[str,pd.DataFrame]:
    """
    Checks if the proposed delta-P / delta-Q at each bus would cause a linear V change more than the limit.
    Returns {table: buses to exclude from scaling, sorted by |dv|}.
    """
    sens_df = pvqv_df[['Number', 'Name', 'NomkV', 'SensdVdPself', 'SensdVdQself']]
    result = {}
    for table, df in targets.items():
        p, q = PVQV_FIELDS[table]
        delta_df = pd.DataFrame({
            'BusNum': df['BusNum']
            ,'dp': df[f'{p}_Target'] - df[p]
            ,'dq': (df[f'{q}_Target'] - df[q]) if q else 0.0
        }).groupby('BusNum', as_index=False).sum()
        bus_df = sens_df.merge(delta_df, left_on='Number', right_on='BusNum', how='inner').drop(columns=['BusNum'])
        bus_df['dv'] = bus_df['SensdVdPself'] * bus_df['dp'] + bus_df['SensdVdQself'] * bus_df['dq']
        result[table] = bus_df[bus_df['dv'].abs() > PVQV_DV_LIMIT].sort_values(by='dv', key=abs, ascending=False)
    return result

def voltage_ok(df: pd.DataFrame, vpu_by_bus: pd.Series) -> pd.Series:
    return df['BusNum'].map(vpu_by_bus).between(EXCLUDE_VMIN, EXCLUDE_VMAX, inclusive='neither')

def bus_vpu(SimAuto) -> pd.Series:
    return pw.get_df(SimAuto, 'Bus', {'Number': int, 'Vpu': float}).set_index('Number')['Vpu']

MAIN_ISLAND = 1     # Bus IslandNumber of the main island. Any other (or blank) is a separate island.

def bus_islands(SimAuto) -> pd.Series:
    """IslandNumber of each bus (NaN where blank), by bus number."""
    return pw.get_df(SimAuto, 'Bus', {'Number': int, 'IslandNumber': int}).set_index('Number')['IslandNumber']

def seed_islands(SimAuto, swing_buses) -> pd.Series | None:
    """
    With the seed PWB open: solves it, and returns bus_islands(), or None (islands not checked, with a warning) if the
    case doesn't solve or no swing bus is in island MAIN_ISLAND (the numbering isn't what this code expects). A live
    second island has a slack bus of its own: one swing bus in the main island is enough.
    """
    if not pw.solve(SimAuto):
        print('WARNING: the seed does not solve; islands are not checked.')
        return None
    islands = bus_islands(SimAuto)
    on_swing = islands.reindex(list(swing_buses))
    if len(on_swing) and not (on_swing == MAIN_ISLAND).any():
        print(f'WARNING: the swing bus is in island {on_swing.tolist()}, not {MAIN_ISLAND}; islands are not checked.')
        return None
    return islands

def island_reason(df: pd.DataFrame, island_by_bus: pd.Series | None) -> pd.Series:
    """
    Per row: '' if its bus is in the main island (or islands weren't checked), else why it is left out:
    'Island 3, not the main island', or 'No island' (blank IslandNumber, or a bus not found).
    """
    if island_by_bus is None or not len(df):
        return pd.Series('', index=df.index, dtype=object)
    island = df['BusNum'].map(island_by_bus)
    reason = island.map(lambda i: 'No island' if pd.isna(i) else f'Island {int(i)}, not the main island')
    return reason.where(island != MAIN_ISLAND, '').astype(object)

def is_island_reason(reason: pd.Series) -> pd.Series:
    return reason.astype(str).str.startswith(('Island ', 'No island'))

def print_islanded(counts: dict[str, int], reasons):
    """One line for the units found outside the main island, if any."""
    if any(counts.values()):
        names = sorted({r.split(',')[0].replace('Island ', '') if r != 'No island' else 'blank' for r in reasons})
        print(f'Outside the main island, to be turned off: {counts.get("Gen", 0)} gens and {counts.get("Load", 0)} loads '
              f'(islands {", ".join(names)}).')

def turn_off_targets(df: pd.DataFrame, table: str, mask):
    """Targets of the rows in mask: every status Open, and the MW/Mvar fields it zeroes at 0 (STATUS)."""
    for status, zeroed in STATUS[table].items():
        df.loc[mask, f'{status}_Target'] = 'Open'
        df.loc[mask, [f'{c}_Target' for c in zeroed]] = 0.0

def exclude_islands(targets, island_by_bus: pd.Series | None):
    """
    Per-unit (topology) targets: gens/loads outside the main island (island_reason()) are excluded from scaling, with
    that reason even over an earlier one (e.g. a second island's own 'Swing Unit'), and targeted Open at 0 MW, for
    turn_off_islanded() to write.
    """
    counts, reasons = {}, set()
    for table, df in targets.items():
        r = island_reason(df, island_by_bus)
        off = (r != '') & (df['ID'] != prep.SWING_ID)     # Never the giant swing.
        df.loc[off, 'ExclusionReason'] = r[off]
        turn_off_targets(df, table, off)
        counts[table] = int((off & (df['Status'] == 'Closed')).sum())
        reasons |= set(r[off])
    print_islanded(counts, reasons)

def turn_off_islanded(SimAuto, targets):
    """
    With the case open, before scaling: writes every gen/load outside the main island (its ExclusionReason is an island
    one) that is Closed now Open at 0 MW, as its targets say. Each write is solve-checked and bisected; those which
    diverge stay as they are (printed; the exceptions sheet's Final Status shows them Closed).
    """
    counts, failed_ids = {}, []
    for table in ('Load', 'Gen'):       # Loads first: a live island's gens (its slack too) can't go while it has load.
        df = targets[table]
        rows = df[is_island_reason(df['ExclusionReason']) & (df['Status'] == 'Closed')]
        if not len(rows):
            continue
        new_df = rows.copy()
        for c in [*STATUS[table], *SCALED[table]]:
            new_df[c] = rows[f'{c}_Target']
        failed = apply_rows(SimAuto, table, df, new_df[['BusNum', 'ID', *STATUS[table], *SCALED[table]]])
        counts[table] = len(rows) - len(failed)
        failed_ids += list(failed)
    if counts:
        print(f'Turned off {counts.get("Gen", 0)} gens and {counts.get("Load", 0)} loads outside the main island'
              + (f'; {len(failed_ids)} stay on (the case diverged without them): {", ".join(map(str, failed_ids[:5]))}'
                 + (' ...' if len(failed_ids) > 5 else '') if failed_ids else '') + '.')

def hold_voltage(targets, vpu_by_bus: pd.Series):
    # Once a bus exceeds the vmin/vmax limits, hold its gens/loads back from scaling until it recovers.
    # This avoids voltage collapse related issues during scaling.
    for df in targets.values():
        # A row already at its target isn't held (it would end excluded, and its status never set). Before mark_start() no row has a Progress.
        hold(df, ~voltage_ok(df, vpu_by_bus) & (df['Progress'] < RESOLUTION if 'Progress' in df else True), 'Voltage')

# ------------------ Applying changes ------------------
def apply_rows(SimAuto, table: str, df: pd.DataFrame, new_df: pd.DataFrame, solves: int = 1) -> list:
    """
    Applies new_df (rows of df, with new values) to the case, bisecting to find rows which diverge.
    Records the rest in df. Returns the index labels of rows which diverged (left unchanged).
    """
    failed = pw.set_df_bisect(SimAuto, table, new_df, write_fields(table), solves)
    ok_df = new_df.drop(index=failed)
    df.loc[ok_df.index, ok_df.columns] = ok_df
    return failed

def apply_or_exclude(SimAuto, table: str, df: pd.DataFrame, new_df: pd.DataFrame, reason: str):
    """Like apply_rows(), but excludes rows which diverge from scaling."""
    failed = apply_rows(SimAuto, table, df, new_df)
    exclude(df, df.index.isin(failed), reason)

def close_all_related_gen_load(SimAuto, targets, label='Main scaling'):
    """
    Closes included gens & loads which are open now but closed in the target, at 0 MW / 0 Mvar.
    Does nothing (and prints nothing) when no status changes, as in a balance.
    """
    for table, df in targets.items():
        new_df = df[included(df)].copy()
        changed = pd.Series(False, index=new_df.index)
        for status, zeroed in STATUS[table].items():
            to_close = (new_df[status] == 'Open') & (new_df[f'{status}_Target'] == 'Closed')
            new_df.loc[to_close, zeroed] = 0.0
            new_df.loc[to_close, status] = 'Closed'
            changed |= to_close
        if changed.any():
            print(f'{label}: closing {changed.sum()} {table.lower()}s in at 0 MW')
        apply_or_exclude(SimAuto, table, df, new_df[changed], 'Diverged closing at 0 MW')

def set_gen_load_status(SimAuto, targets, label='Main scaling'):
    """
    Sets included gens & loads to their final target statuses. Rows which diverge are excluded.
    Only rows whose status changes are written, so a balance (no status changes) writes and solves nothing.
    """
    # Loads first: any being opened should have reached 0 MW / 0 Mvar, so opening them has no impact on the solution.
    # Then gens: opening them at 0 MW may still remove VAR support. If one cannot be opened, make note of it and proceed.
    for table in ('Load', 'Gen'):
        df = targets[table]
        new_df = df[included(df)]
        new_df = new_df.assign(**{s: new_df[f'{s}_Target'] for s in STATUS[table]})
        changed = pd.concat([new_df[s] != df.loc[new_df.index, s] for s in STATUS[table]], axis=1).any(axis=1)
        new_df = new_df[changed]
        if len(new_df):
            print(f'{label}: {len(new_df)} {table.lower()}s change status')
        apply_or_exclude(SimAuto, table, df, new_df, 'Diverged setting final status')

def mark_start(targets):
    """Records where scaling starts from. Call after close_all_related_gen_load(), which changes MW/Mvar."""
    for table, df in targets.items():
        for c in SCALED[table]:
            df[f'{c}_Start'] = df[c]
        df['Progress'] = 0

def at_progress(table: str, df: pd.DataFrame, progress) -> pd.DataFrame:
    """Rows of df, with scaled fields moved to `progress` (in 1/RESOLUTION units) of the way from start to target."""
    new_df = df.assign(Progress=progress)
    for c in SCALED[table]:
        new_df[c] = df[f'{c}_Start'] + (df[f'{c}_Target'] - df[f'{c}_Start']) * new_df['Progress'] / RESOLUTION
    return new_df

def rescale(df: pd.DataFrame, factor: pd.Series, cols: list[str], cap: pd.Series | None = None):
    """
    After rows were moved by `factor` (indexed by ObjectID) in the middle of scaling: multiplies their current, start and
    target values of `cols` in df by it, so at_progress() carries on from where they are instead of undoing the move.
    cap: optional upper limit (e.g. MWMax) for the targets. A target already above it isn't raised.
    """
    f = factor[factor.index.isin(df.index)]
    for c in cols:
        for name in (c, f'{c}_Start', f'{c}_Target'):
            if name not in df.columns:
                continue
            new = df.loc[f.index, name] * f
            if cap is not None and name == f'{c}_Target':
                old, limit = df.loc[f.index, name], cap.reindex(f.index)
                new = new.clip(upper=limit.where(limit > old, old))
            df.loc[f.index, name] = new

# ------------------ Voltage support ------------------
def create_statcom_on_lowestv_bus(SimAuto, skip_buses, label='Main scaling'):
    """
    If there is a bus with voltage lower than STATCOM_VPU, adds a STATCOM to the lowest voltage bus.
    Returns None if no bus needs one, otherwise (bus number, whether the STATCOM closed in successfully).
    Buses which already have one in the case (from any scaling run, e.g. an earlier balance) are skipped, as are
    skip_buses (this run's attempts, including those rolled back).
    """
    gen_df = pw.get_df(SimAuto, 'Gen', {'BusNum': int, 'ID': str})
    skip_buses = set(skip_buses) | set(gen_df.loc[gen_df['ID'] == STATCOM_ID, 'BusNum'])
    bus_fields: dict[str,type] = {
        'Number': int
        ,'BusNomVolt': float
        ,'Vpu': float
        ,'IslandNumber': int
        ,'BusIsStarBus:1': str
    }
    bus_df = pw.get_df(SimAuto, 'Bus', bus_fields)
    filtered_df = bus_df[(bus_df['IslandNumber'] == 1) &
                         (bus_df['BusIsStarBus:1'] == 'NO') &
                         (bus_df['BusNomVolt'] > STATCOM_VNOM_MIN) &
                         (bus_df['Vpu'] < STATCOM_VPU) &
                         ~bus_df['Number'].isin(skip_buses)]
    if len(filtered_df) == 0:
        return None

    min_row = filtered_df.loc[filtered_df['Vpu'].idxmin()]
    number = int(min_row['Number'])

    # Create a STATCOM, Closed. Roll back if it doesn't solve.
    statcom_dict = {
        'BusNum': number
        ,'GenID': STATCOM_ID
        ,'GenStatus': 'Closed'
        ,'GenAVRAble': 'YES'
        ,'GenMVRMax': 9999
        ,'GenMVRMin': -9999
        ,'GenMvrSetPoint': 0
        ,'GenAGCAble': 'NO'
        ,'GenMWMax': 0
        ,'GenMWMin': 0
        ,'GenMWSetPoint': 0
        ,'GenVoltSet': 1.000
    }
    statcom_df = pd.DataFrame([statcom_dict])

    with pw.trial(SimAuto) as t:
        with pw.creating(SimAuto):
            pw.set_df(SimAuto, 'Gen', statcom_df, list(statcom_dict))
    print(f'{label}: STATCOM on bus {number} ({min_row["Vpu"]:.2f} pu): {"closed" if t.ok else "rolled back"}')

    return (number, t.ok)

def drop_collapsed_sections(SimAuto, label='Main scaling') -> list[str]:
    """
    Disconnects network sections which are beginning to show collapse.
    I.e. opens branches connecting from buses with OK voltage (>COLLAPSE_VPU) to buses with awful voltage (<COLLAPSE_VPU).
    Returns the ObjectIDs of branches dropped.
    """
    df = pw.get_df(SimAuto, 'Branch', {'ObjectID': str, 'Status': str, 'BranchVpuHigh': float, 'BranchVpuLow': float})
    filtered_df = df[(df['Status'] == 'Closed') & (df['BranchVpuHigh'] > COLLAPSE_VPU) & (df['BranchVpuLow'] < COLLAPSE_VPU)]
    if len(filtered_df) == 0:
        return []

    dropped = filtered_df['ObjectID'].unique().tolist()
    with pw.trial(SimAuto) as t:
        pw.set_df(SimAuto, 'Branch', filtered_df.assign(Status='Open'), ['ObjectID', 'Status'])
        result = SimAuto.RunScriptCommand('ClearSmallIslands;')
        if result[0]:
            print(f'WARNING: ClearSmallIslands: {result[0]}')
    if not t.ok:
        print(f'{label}: opening {len(dropped)} branches into collapsed sections: rolled back')
        return []
    print(f'{label}: opened {len(dropped)} branches into collapsed sections')
    return dropped

# ------------------ Scaling ------------------
def catch_up(SimAuto, targets, progress: int, step: int, bisect: bool) -> int:
    """
    Retries held gens/loads on buses within the voltage band, moving each up to `step` closer to `progress`.
    Those which reach `progress` resume scaling with the rest.
    bisect=False: one attempt with all of them together. bisect=True: bisect to find those which still diverge.
    Returns how many moved.
    """
    vpu_by_bus = bus_vpu(SimAuto)
    moved = 0
    for table, df in targets.items():
        rows = df[held(df) & voltage_ok(df, vpu_by_bus)]
        new_df = at_progress(table, rows, (rows['Progress'] + step).clip(upper=progress))
        new_df = new_df[new_df['Progress'] > rows['Progress']]
        if len(new_df) == 0:
            failed = []
        elif bisect:
            failed = apply_rows(SimAuto, table, df, new_df, STEP_SOLVES)
        else:
            with pw.trial(SimAuto, solves=STEP_SOLVES) as t:
                pw.set_df(SimAuto, table, new_df, write_fields(table))
            failed = [] if t.ok else list(new_df.index)
            if t.ok:
                df.loc[new_df.index, new_df.columns] = new_df
        moved += len(new_df) - len(failed)
        df.loc[rows.index[df.loc[rows.index, 'Progress'] >= progress], 'HeldReason'] = ''
    return moved

def retry_held(SimAuto, targets, label: str):
    """
    Once the rest have reached their targets: retries held gens/loads, bisecting to find those which still diverge.
    Tries one jump to target first, and halves the jump whenever a round moves nothing.
    Those still held at the end are excluded, left where they got to.
    """
    step = MAX_STEP
    for _ in range(FINAL_MAX_ROUNDS):
        vpu_by_bus = bus_vpu(SimAuto)
        hold_voltage(targets, vpu_by_bus)
        # Only rows on buses inside the voltage band can be retried; with none, halving the step changes nothing.
        if not any((held(df) & voltage_ok(df, vpu_by_bus)).any() for df in targets.values()):
            break
        print(f'{label}: retrying held gens/loads, {step}/{RESOLUTION} at a time.')
        if catch_up(SimAuto, targets, RESOLUTION, step, bisect=True):
            prep.adjust_shunts(SimAuto)
        elif step > 1:
            step //= 2
        else:
            break
    for df in targets.values():
        mask = held(df)
        df.loc[mask, 'ExclusionReason'] = df.loc[mask, 'HeldReason']

# ------------------ Run collection ------------------
@dataclass
class Run:
    """
    One scaling run, for the `exceptions` and `runs` sheets.
    name: 'prebalance', 'main', 'midrun<N>', 'distgen', 'makeup', 'pretest' or 'balance<N>'. log: what iterate_to_gen_load_targets() returned (or the error's `.log`).
    moved: a balance's balance_targets() tables, after scaling (None for the main run). before / after: the swing's MW
    around the run (NaN: not measured). at: a midrun balance's main progress, e.g. '50 of 100'. note: short flags.
    """
    name: str
    log: dict = field(default_factory=dict)
    moved: dict | None = None
    before: float = np.nan
    after: float = np.nan
    at: str = ''
    note: str = ''

EXCEPTION_COLUMNS = ['Table', 'ObjectID', 'AreaNumber', 'Group', 'BusNum', 'Seed MW', 'Target MW', 'Final MW', 'Final Status', 'Reason']
RUN_COLUMNS = ['Run', 'At', 'Swing MW Before', 'Swing MW After', 'Gen MW Moved', 'Load MW Moved', 'Steps', 'Diverged Steps', 'Note']
_P = ['SMW', 'IMW', 'ZMW']

def _by_id(df: pd.DataFrame) -> pd.DataFrame:
    return df.set_index('ObjectID') if 'ObjectID' in df.columns else df

def _mw(table: str, df: pd.DataFrame, suffix: str = '') -> pd.Series:
    """Gen MWSetPoint, or Load SMW + IMW + ZMW (the columns present), each with `suffix`."""
    if table == 'Gen':
        return df[f'MWSetPoint{suffix}']
    return df[[f'{c}{suffix}' for c in _P if f'{c}{suffix}' in df.columns]].sum(axis=1)

def _element_info(table: str, df: pd.DataFrame) -> dict:
    """{ObjectID: AreaNumber, Group, BusNum, Seed MW, Target MW} of a targets table (blank where the table lacks it)."""
    if not len(df):
        return {}
    from . import area      # Not at import: area imports this module.
    nan = pd.Series(np.nan, index=df.index)
    if table == 'Gen':
        seed = df['MWSetPoint_Seed'] if 'MWSetPoint_Seed' in df else nan
        group = area.classify_gen_type(df) if {'UnitTypeCode', 'FuelTypeCode'} <= set(df.columns) else pd.Series('', index=df.index)
        target = df['Planned MW'] if 'Planned MW' in df else df['MWSetPoint_Target'] if 'MWSetPoint_Target' in df else nan
    else:
        seed = df['P_Seed'] if 'P_Seed' in df else _mw(table, df, '_Seed') if f'{_P[0]}_Seed' in df else nan
        group = pd.Series('Load', index=df.index)
        target = df['Planned MW'] if 'Planned MW' in df else _mw(table, df, '_Target') if f'{_P[0]}_Target' in df else nan
    frame = pd.DataFrame({'AreaNumber': df['AreaNumber'] if 'AreaNumber' in df else nan, 'Group': group,
                          'BusNum': df['BusNum'] if 'BusNum' in df else nan, 'Seed MW': seed, 'Target MW': target})
    return frame.to_dict('index')

def exceptions(runs: list[Run], targets: dict[str, pd.DataFrame] | None = None, final: dict[str, pd.DataFrame] | None = None) -> pd.DataFrame:
    """
    The `exceptions` sheet (EXCEPTION_COLUMNS): a row per gen, load, STATCOM or opened branch not left on plan.
    Reason: '<run>: <reason>' for each run which excluded it or changed the network, joined by '; ' in run order.
    Gens/loads whose final status differs from their target (main `targets`) with no other reason get 'not on target status'.
    The swing ('Swing Unit') is left out. targets: the main scaling's tables (Seed, Target; None: blank). final: read_case()
    tables of the scaled case, with ObjectID, Status and MW (None: Final columns blank).
    """
    reasons: dict[tuple, list[str]] = {}
    network: dict[tuple, dict] = {}
    for run in runs:
        for table in ('Gen', 'Load'):
            df = run.log.get(table.lower())
            if df is None or not len(df):
                continue
            for oid, reason in df['ExclusionReason'].items():
                if reason and reason != 'Swing Unit':
                    reasons.setdefault((table, oid), []).append(f'{run.name}: {reason}')
        sc = run.log.get('statcom_bus_df')
        if sc is not None and 'BusNum' in sc.columns:
            for bus, closed in zip(sc['BusNum'], sc['Closed']):
                key = ('STATCOM', f'{bus} {STATCOM_ID}')
                reasons.setdefault(key, []).append(f'{run.name}: ' + ('STATCOM added' if closed else 'STATCOM tried, rolled back (not in the case)'))
                network[key] = {'BusNum': bus, 'Final Status': 'Closed' if closed else ''}
        br = run.log.get('dropped_branch_df')
        if br is not None and 'ObjectID' in br.columns:
            for oid in br['ObjectID']:
                reasons.setdefault(('Branch', oid), []).append(f'{run.name}: branch opened (collapse)')
                network[('Branch', oid)] = {'Final Status': 'Open'}

    finals = {t: _by_id(df) for t, df in (final or {}).items()}
    info: dict[str, dict] = {}
    for table in ('Gen', 'Load'):
        sources = [] if targets is None or table not in targets else [targets[table]]
        sources += [r.log[table.lower()] for r in runs if table.lower() in r.log]
        info[table] = {}
        for df in sources:
            for oid, v in _element_info(table, df).items():
                info[table].setdefault(oid, v)
        if targets is not None and table in targets and table in finals and 'Status_Target' in targets[table]:
            t, f = targets[table], finals[table]
            common = t.index.intersection(f.index)
            off = (t.loc[common, 'Status_Target'] != f.loc[common, 'Status']) & (t.loc[common, 'ExclusionReason'] != 'Swing Unit')
            for oid in common[off.to_numpy()]:
                reasons.setdefault((table, oid), ['not on target status'])

    rows = []
    for key in sorted(reasons, key=lambda k: ('Gen', 'Load', 'STATCOM', 'Branch').index(k[0])):
        table, oid = key
        row = {'Table': table, 'ObjectID': oid, 'AreaNumber': pd.NA, 'Group': '', 'BusNum': pd.NA, 'Seed MW': np.nan,
               'Target MW': np.nan, 'Final MW': np.nan, 'Final Status': '', **network.get(key, {}), **info.get(table, {}).get(oid, {})}
        if table in finals and oid in finals[table].index:
            f = finals[table].loc[[oid]]
            row['Final MW'], row['Final Status'] = float(_mw(table, f).iloc[0]), f['Status'].iloc[0]
        row['Reason'] = '; '.join(dict.fromkeys(reasons[key]))
        rows.append(row)
    out = pd.DataFrame(rows, columns=EXCEPTION_COLUMNS)
    for c in ('AreaNumber', 'BusNum'):
        out[c] = pd.to_numeric(out[c]).astype('Int64')
    for c in ('Seed MW', 'Target MW', 'Final MW'):
        out[c] = out[c].astype(float)
    return out

def runs_sheet(runs: list[Run]) -> pd.DataFrame:
    """The `runs` sheet (RUN_COLUMNS): a row per scaling run. MW moved: balances only (area.mw_changes())."""
    from . import area
    rows = []
    for run in runs:
        moved = area.mw_changes(run.moved) if run.moved is not None else {}
        steps = run.log.get('step_df')
        has = steps is not None and 'Result' in steps.columns
        rows.append({'Run': run.name, 'At': run.at, 'Swing MW Before': run.before, 'Swing MW After': run.after,
                     'Gen MW Moved': moved.get('Gen MW Change', np.nan), 'Load MW Moved': moved.get('Load MW Change', np.nan),
                     'Steps': len(steps) if has else 0, 'Diverged Steps': int((steps['Result'] != 'Solved').sum()) if has else 0,
                     'Note': run.note})
    out = pd.DataFrame(rows, columns=RUN_COLUMNS)
    for c in ('Swing MW Before', 'Swing MW After', 'Gen MW Moved', 'Load MW Moved'):
        out[c] = out[c].astype(float)
    return out

def iterate_to_gen_load_targets(SimAuto, targets, pvqv_df=None, after_step=None, label='Main scaling',
                                initial_step=None) -> dict[str,pd.DataFrame]:
    """
    Scales included gens & loads from their start to their targets. Returns the scale log sheets.
    Steps are in 1/RESOLUTION units. Each step which solves at its first size doubles the next one (up to MAX_STEP;
    one which solved only once halved keeps its size),
    and each which diverges is retried at half the size. A 1/RESOLUTION step which still diverges is bisected:
    the gens/loads which diverge are held back at their last good values, and the rest go ahead.
    Held gens/loads (including those on buses outside the voltage band) are retried after each step, and again at the end.
    pvqv_df: optional PVQV sensitivities. If given, gens/loads with a large linear dV estimate are excluded up front.
    after_step: optional function(SimAuto, targets, progress), called after each step (e.g. area.midrun_balancer()).
        If it returns a number, the next step is at most that (at least 1), e.g. to keep the swing from moving too far.
    label: names this run in the printed progress, to tell a balance's own scaling from the main one.
    initial_step: first step (default INITIAL_STEP). Balances, a move of a few percent, start at MAX_STEP.
    """
    if not pw.solve(SimAuto):
        raise pw.SimAutoError('Could not solve the original input case.')

    for df in targets.values():
        df['HeldReason'] = ''
    pvqv = {}
    if pvqv_df is not None:
        pvqv = compute_pvqv_exclusions(targets, pvqv_df)
        pvqv_buses = pd.concat([bus_df['Number'] for bus_df in pvqv.values()])
        for df in targets.values():
            exclude(df, df['BusNum'].isin(pvqv_buses), 'PVQV')
    hold_voltage(targets, bus_vpu(SimAuto))
    close_all_related_gen_load(SimAuto, targets, label)
    mark_start(targets)

    prep.adjust_shunts(SimAuto)     # Leaves the case solved: each change is kept only if it solves.

    progress, step = 0, initial_step or INITIAL_STEP
    failed_step = None              # The step size which first diverged, before halving.
    step_log = []
    statcoms: dict[int,bool] = {}   # Bus number -> whether the STATCOM closed in successfully.
    dropped_branches: list[str] = []
    def log(start, goal, result, caught_up=0):
        held_count = sum(int(held(df).sum()) for df in targets.values())
        step_log.append({'From': start, 'To': goal, 'Result': result, 'CaughtUp': caught_up, 'Held': held_count})

    def package():
        return {
            **{table.lower(): df[~included(df)] for table, df in targets.items()}
            ,'step_df': pd.DataFrame(step_log)
            ,'statcom_bus_df': pd.DataFrame({"BusNum": list(statcoms), "Closed": list(statcoms.values())})
            ,'dropped_branch_df': pd.DataFrame({"ObjectID": dropped_branches})
        }

    try:
        while progress < RESOLUTION:
            step = min(step, RESOLUTION - progress)
            goal = progress + step
            new = {table: at_progress(table, df[scaling(df)], goal) for table, df in targets.items()}
            with pw.trial(SimAuto, solves=STEP_SOLVES) as t:
                for table, new_df in new.items():
                    pw.set_df(SimAuto, table, new_df, write_fields(table))
            if t.ok:
                for table, new_df in new.items():
                    targets[table].loc[new_df.index, new_df.columns] = new_df
                result = 'Solved'
            elif step > 1:
                print(f'{label}: {progress} -> {goal} of {RESOLUTION}: diverged, halving')
                log(progress, goal, 'Diverged, halved step')
                failed_step = failed_step or step
                step //= 2
                continue
            else:
                for table, new_df in new.items():
                    df = targets[table]
                    hold(df, df.index.isin(apply_rows(SimAuto, table, df, new_df, STEP_SOLVES)), 'Diverged during scaling')
                result = 'Bisected'
            start, progress = progress, goal

            prep.adjust_shunts(SimAuto)
            hold_voltage(targets, bus_vpu(SimAuto))
            held_now = sum(int(held(df).sum()) for df in targets.values())
            outcome = result.lower() + (f' ({held_now} held)' if held_now else '')
            print(f'{label}: {start} -> {goal} of {RESOLUTION}: {outcome}')
            if DROP_COLLAPSED:
                dropped_branches += drop_collapsed_sections(SimAuto, label)
            statcom = ADD_STATCOMS and create_statcom_on_lowestv_bus(SimAuto, skip_buses=list(statcoms), label=label)
            if statcom:
                statcoms[statcom[0]] = statcom[1]
            log(start, goal, result, catch_up(SimAuto, targets, progress, step, bisect=False))
            cap = after_step(SimAuto, targets, progress) if after_step else None
            if t.ok and not failed_step:
                step = min(step * 2, MAX_STEP)
            elif t.ok:
                pass                    # Solved only once halved: keep this size (doubling back would fail again).
            elif failed_step:
                step = failed_step      # The gens/loads which diverged are held now, so go back to the larger step.
            failed_step = None
            if cap is not None:
                step = min(step, max(1, int(cap)))

        retry_held(SimAuto, targets, label)
        set_gen_load_status(SimAuto, targets, label)
    except pw.SimAutoError as e:
        # What this run did so far (STATCOMs, dropped branches) stays in the case: hand its log to the caller.
        e.log = package()
        raise

    print(f'{label}: done')

    return package()
