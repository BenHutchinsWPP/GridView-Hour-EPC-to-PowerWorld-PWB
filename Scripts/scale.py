"""
Step 02: scale the TopoSeed case's gens & loads to a GridView hour's targets.

Gens and loads are handled by the same code. Their working tables ("targets") are indexed by ObjectID and hold:
    <field>          current value in the case
    <field>_Target   value in the GridView hour
    <field>_Seed     value in the TopoSeed case
    <field>_Delta    change per scaling step
    ExclusionReason  '' while the row is scaled. Otherwise, why it was excluded.
"""
from pathlib import Path
import multiprocessing as mp
import pandas as pd
from . import pw, topo

# ------------------ Tuning thresholds ------------------
EXCLUDE_VMIN = 0.88             # Stop scaling gens/loads on buses below this Vpu...
EXCLUDE_VMAX = 1.12             # ...or above this Vpu.
STATCOM_VPU = 0.85              # Add a STATCOM to the lowest-voltage bus below this Vpu...
STATCOM_VNOM_MIN = 50           # ...with nominal kV above this.
COLLAPSE_VPU = 0.80             # Drop branches between buses above and below this Vpu.
PVQV_DV_LIMIT = 0.10            # Exclude buses whose linear PVQV dV estimate exceeds this.
SCALING_ITERATIONS = 100        # Steps used to scale gens/loads from seed to target.
MAX_WORKERS = None              # Parallel PowerWorld instances for gen target tests. None = CPU count.

STATCOM_ID = 'SC'   # ID of STATCOMs added during scaling.

# ------------------ Gen & load definitions ------------------
# Per table: each status field, and the MW/Mvar fields which must be 0 while it is Open.
STATUS = {
    'Gen': {'Status': ['MWSetPoint']},
    'Load': {'Status': ['SMW', 'SMvar', 'DistMWInput', 'DistMvarInput'], 'DistStatus': ['DistMWInput', 'DistMvarInput']},
}
# The fields which get scaled.
SCALED = {table: statuses['Status'] for table, statuses in STATUS.items()}
# The (MW, Mvar) fields used to estimate voltage change from PVQV sensitivities.
PVQV_FIELDS = {'Gen': ('MWSetPoint', None), 'Load': ('SMW', 'SMvar')}

TARGET_FIELDS = {
    table: {'ObjectID': str, 'BusNum': int, 'BusName': str, 'NomkV': float, 'ID': str,
            **{s: str for s in STATUS[table]}, **{c: float for c in SCALED[table]}}
    for table in STATUS
}

def write_fields(table: str) -> list[str]:
    """The fields written back to PowerWorld."""
    return ['BusNum', 'ID', *STATUS[table], *SCALED[table]]

def included(df: pd.DataFrame) -> pd.Series:
    return df['ExclusionReason'] == ''

def exclude(df: pd.DataFrame, mask, reason: str):
    """Excludes rows from scaling. The first recorded reason is kept."""
    df.loc[pd.Series(mask, index=df.index) & included(df), 'ExclusionReason'] = reason

# ------------------ Targets ------------------
def compute_pw_targets(SimAuto, gv_fp: Path, pw_fp: Path, dummy_bus_fp: Path) -> dict[str,pd.DataFrame]:
    """Returns {'Gen': df, 'Load': df} targets, taking the GridView EPC's gens & loads into the PWB case. See build_targets()."""
    topo.open_gv_case(SimAuto, gv_fp, dummy_bus_fp)
    gv = {table: pw.get_df(SimAuto, table, fields) for table, fields in TARGET_FIELDS.items()}

    pw.open_case(SimAuto, pw_fp)
    case = {table: pw.get_df(SimAuto, table, fields) for table, fields in TARGET_FIELDS.items()}
    SimAuto.CloseCase()

    return build_targets(gv, case)

def build_targets(gv: dict[str,pd.DataFrame], case: dict[str,pd.DataFrame]) -> dict[str,pd.DataFrame]:
    """Puts the case's gens & loads side by side with their GridView targets."""
    targets = {}
    for table, statuses in STATUS.items():
        gv_df = gv[table]
        if table == 'Load':
            # GridView leaves distributed generation in the load table at seed-case values, unused. Target Open & 0 MW.
            gv_df = gv_df.assign(DistStatus='Open', DistMWInput=0.0, DistMvarInput=0.0)

        fields = [*statuses, *SCALED[table]]
        df = case[table].merge(gv_df[['ObjectID', *fields]], on='ObjectID', how='left', suffixes=('', '_Target'))
        for f in fields:
            df[f'{f}_Seed'] = df[f]

        # Elements which aren't in the GridView hour are opened.
        df = df.fillna({**{f'{s}_Target': 'Open' for s in statuses}, **{f'{c}_Target': 0.0 for c in SCALED[table]}})

        # Anything open, now or in the target, is at 0 MW / 0 Mvar.
        for suffix in ('', '_Target'):
            for status, zeroed in statuses.items():
                df.loc[df[status + suffix] == 'Open', [c + suffix for c in zeroed]] = 0.0

        df['ExclusionReason'] = ''
        targets[table] = df.set_index('ObjectID')
    return targets

def test_gen_targets(pw_fp: Path, gen_df: pd.DataFrame) -> pd.DataFrame:
    """
    Tests each generator target individually, and adds a 'Success' column reporting which individual changes solve.
    """
    import win32com.client
    SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")
    pw.open_case(SimAuto, pw_fp)
    if not pw.solve(SimAuto):
        raise pw.SimAutoError(f'Base case does not solve, cannot test gen targets: {pw_fp}')

    target_df = gen_df.assign(Status=gen_df['Status_Target'], MWSetPoint=gen_df['MWSetPoint_Target'])
    success = []
    for i in range(len(target_df)):
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
    """Prints total gen, dist gen & load MW for the base and target cases, and warns if the target imbalance is high."""
    gen_df, load_df = targets['Gen'], targets['Load']
    def totals(suffix):
        gen_mw = gen_df['MWSetPoint' + suffix].sum()
        dist_mw = load_df['DistMWInput' + suffix].sum()
        load_mw = load_df['SMW' + suffix].sum()
        return [gen_mw, dist_mw, load_mw, gen_mw + dist_mw - load_mw]

    total_delta_df = pd.DataFrame(
        {'Base Case': totals(''), 'Target Case': totals('_Target')}
        ,index=['Gen MW', 'Dist Gen MW', 'Load MW', 'Gen + Dist Gen - Load']
    )
    print(total_delta_df.round(0))

    base_total_delta, target_total_delta = total_delta_df.iloc[-1]
    if abs(target_total_delta) > 1.5 * abs(base_total_delta):
        print('-----------------------------------------------------------------------------------')
        print('WARNING: Gen/Load imbalance in the target case is high. This may cause instability.')
        print('-----------------------------------------------------------------------------------')

    return total_delta_df

# ------------------ Exclusions ------------------
def compute_pvqv_exclusions(targets, pvqv_df, delta_v_limit = PVQV_DV_LIMIT) -> dict[str,pd.DataFrame]:
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
        result[table] = bus_df[bus_df['dv'].abs() > delta_v_limit].sort_values(by='dv', key=abs, ascending=False)
    return result

def voltage_ok(df: pd.DataFrame, vpu_by_bus: pd.Series, v_min = EXCLUDE_VMIN, v_max = EXCLUDE_VMAX) -> pd.Series:
    return df['BusNum'].map(vpu_by_bus).between(v_min, v_max, inclusive='neither')

def compute_voltage_exclusions(SimAuto, targets, v_min = EXCLUDE_VMIN, v_max = EXCLUDE_VMAX):
    # Once a bus exceeds the vmin/vmax limits, take it out of the scaling equation.
    # This ensures no voltage collapse related issues occur during scaling.
    vpu_by_bus = pw.get_df(SimAuto, 'Bus', {'Number': int, 'Vpu': float}).set_index('Number')['Vpu']
    for df in targets.values():
        exclude(df, ~voltage_ok(df, vpu_by_bus, v_min, v_max), 'Voltage')

# ------------------ Applying changes ------------------
def apply_or_exclude(SimAuto, table: str, df: pd.DataFrame, new_df: pd.DataFrame, reason: str):
    """
    Applies new_df (rows of df, with new values) to the case, then records those values in df.
    Rows which diverge are left unchanged in df, and excluded from scaling.
    """
    fields = write_fields(table)
    failed = pw.set_df_bisect(SimAuto, table, new_df, fields)
    ok_df = new_df.drop(index=failed)
    df.loc[ok_df.index, fields] = ok_df[fields]
    exclude(df, df.index.isin(failed), reason)

def close_all_related_gen_load(SimAuto, targets):
    """Closes included gens & loads which are open now but closed in the target, at 0 MW / 0 Mvar."""
    print('close_all_related_gen_load()')
    for table, df in targets.items():
        new_df = df[included(df)].copy()
        changed = pd.Series(False, index=new_df.index)
        for status, zeroed in STATUS[table].items():
            to_close = (new_df[status] == 'Open') & (new_df[f'{status}_Target'] == 'Closed')
            new_df.loc[to_close, zeroed] = 0.0
            new_df.loc[to_close, status] = 'Closed'
            changed |= to_close
        apply_or_exclude(SimAuto, table, df, new_df[changed], 'Diverged closing at 0 MW')

def set_gen_load_status(SimAuto, targets):
    """Sets included gens & loads to their final target statuses. Rows which diverge are excluded."""
    print('set_gen_load_status()')
    # Loads first: any being opened should have reached 0 MW / 0 Mvar, so opening them has no impact on the solution.
    # Then gens: opening them at 0 MW may still remove VAR support. If one cannot be opened, make note of it and proceed.
    for table in ('Load', 'Gen'):
        df = targets[table]
        new_df = df[included(df)]
        new_df = new_df.assign(**{s: new_df[f'{s}_Target'] for s in STATUS[table]})
        apply_or_exclude(SimAuto, table, df, new_df, 'Diverged setting final status')

def compute_deltas(targets, iterations):
    for table, df in targets.items():
        for c in SCALED[table]:
            df[f'{c}_Delta'] = (df[f'{c}_Target'] - df[c]) / iterations

def next_step(table: str, df: pd.DataFrame) -> pd.DataFrame:
    """Included rows, moved one step towards their targets."""
    step_df = df[included(df)].copy()
    for c in SCALED[table]:
        step_df[c] += step_df[f'{c}_Delta']
    return step_df

# ------------------ Voltage support ------------------
def create_statcom_on_lowestv_bus(SimAuto, skip_buses, vpu_min = STATCOM_VPU, vnom_min = STATCOM_VNOM_MIN):
    """
    If there is a bus with voltage lower than vpu_min, adds a STATCOM to the lowest voltage bus.
    Returns None if no bus needs one, otherwise (bus number, whether the STATCOM closed in successfully).
    """
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
                         (bus_df['BusNomVolt'] > vnom_min) &
                         (bus_df['Vpu'] < vpu_min) &
                         ~bus_df['Number'].isin(skip_buses)]
    if len(filtered_df) == 0:
        return None

    min_row = filtered_df.loc[filtered_df['Vpu'].idxmin()]
    number = int(min_row['Number'])
    print(f'Adding statcom to: {number}, kV={min_row["BusNomVolt"]}, Vpu={min_row["Vpu"]}')

    # Create a STATCOM, in the "Open" position. Solve case.
    statcom_dict = {
        'BusNum': number
        ,'GenID': STATCOM_ID
        ,'GenStatus': 'Open'
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
    if not t.ok:
        print('WARNING: Did not solve after creating statcom in Open state. Rolled back change.')
        return (number, False)

    # Try to close STATCOM. Roll back if it doesn't solve.
    with pw.trial(SimAuto) as t:
        pw.set_df(SimAuto, 'Gen', statcom_df.assign(GenStatus='Closed'), ['BusNum', 'GenID', 'GenStatus'])
    if not t.ok:
        print('WARNING: Did not solve after closing statcom. Rolled back change.')

    return (number, t.ok)

def drop_collapsed_sections(SimAuto, vpu = COLLAPSE_VPU) -> list[str]:
    """
    Disconnects network sections which are beginning to show collapse.
    I.e. opens branches connecting from buses with OK voltage (>vpu) to buses with awful voltage (<vpu).
    Returns the ObjectIDs of branches dropped.
    """
    df = pw.get_df(SimAuto, 'Branch', {'ObjectID': str, 'Status': str, 'BranchVpuHigh': float, 'BranchVpuLow': float})
    filtered_df = df[(df['Status'] == 'Closed') & (df['BranchVpuHigh'] > vpu) & (df['BranchVpuLow'] < vpu)]
    if len(filtered_df) == 0:
        return []

    dropped = filtered_df['ObjectID'].unique().tolist()
    print(f'Dropping branches to collapsed network sections: {dropped}')
    with pw.trial(SimAuto) as t:
        pw.set_df(SimAuto, 'Branch', filtered_df.assign(Status='Open'), ['ObjectID', 'Status'])
        pw.run(SimAuto, 'ClearSmallIslands;')
    if not t.ok:
        print('WARNING: Did not solve after dropping branches. Rolled back change.')
        return []
    return dropped

# ------------------ Scaling ------------------
def iterate_to_gen_load_targets(SimAuto, targets, pvqv_df, iterations=SCALING_ITERATIONS, bisect=False) -> dict[str,pd.DataFrame]:
    """
    Scales included gens & loads to their targets in `iterations` steps. Returns the scale log sheets.
    If a step diverges:
        bisect=False: stop scaling there.
        bisect=True: bisect the step to find the gens/loads which diverge, exclude them at their last good step, and continue.
    """
    if not pw.solve(SimAuto):
        print('Could not solve the original input case!')
        return {'iteration_df': pd.DataFrame({"Value": ['Failed to converge base case.']})}

    pvqv = compute_pvqv_exclusions(targets, pvqv_df)
    pvqv_buses = pd.concat([bus_df['Number'] for bus_df in pvqv.values()])
    for df in targets.values():
        exclude(df, df['BusNum'].isin(pvqv_buses), 'PVQV')
    compute_voltage_exclusions(SimAuto, targets)
    close_all_related_gen_load(SimAuto, targets)
    compute_deltas(targets, iterations)

    topo.adjust_shunts(SimAuto)
    if not pw.solve(SimAuto):
        print('Could not solve after adjusting shunts in the original case!')
        return {'iteration_df': pd.DataFrame({"Value": ['Failed to converge base case with shunt adjustments.']})}

    completed = 0
    statcoms: dict[int,bool] = {}   # Bus number -> whether the STATCOM closed in successfully.
    dropped_branches: list[str] = []
    for iteration in range(1, iterations + 1):
        print(f'----- Iteration: {iteration} of {iterations} -----')
        steps = {table: next_step(table, df) for table, df in targets.items()}
        with pw.trial(SimAuto, solves=2) as t:
            for table, step_df in steps.items():
                pw.set_df(SimAuto, table, step_df, write_fields(table))
        if t.ok:
            for table, step_df in steps.items():
                targets[table].loc[step_df.index, SCALED[table]] = step_df[SCALED[table]]
        elif bisect:
            print(f'Iteration {iteration} did not solve. Bisecting to find the gens/loads which diverge.')
            for table, step_df in steps.items():
                apply_or_exclude(SimAuto, table, targets[table], step_df, 'Diverged during scaling')
        else:
            print(f'Iteration {iteration} did not solve. Reverted iteration and stopping.')
            break
        completed = iteration

        topo.adjust_shunts(SimAuto)
        compute_voltage_exclusions(SimAuto, targets)
        dropped_branches += drop_collapsed_sections(SimAuto)
        statcom = create_statcom_on_lowestv_bus(SimAuto, skip_buses=list(statcoms))
        if statcom:
            statcoms[statcom[0]] = statcom[1]

    if completed == iterations:
        compute_voltage_exclusions(SimAuto, targets)
        set_gen_load_status(SimAuto, targets)

    print('Reached end of iterate_to_gen_load_targets(). Returning.')

    # Package the logs for return.
    return {
        **{table.lower(): df[~included(df)] for table, df in targets.items()}
        ,'iteration_df': pd.DataFrame({"Value": [f'Completed {completed} of {iterations} iterations']})
        ,'statcom_bus_df': pd.DataFrame({"BusNum": list(statcoms), "Closed": list(statcoms.values())})
        ,**{f'{table.lower()}_pvqv': bus_df for table, bus_df in pvqv.items()}
        ,'dropped_branch_df': pd.DataFrame({"ObjectID": dropped_branches})
    }
