"""
Step 02 targets: scale the PWB to the GridView hour by area, not element by element.

Topology is not matched: nothing is created or deleted, and no branch or load is opened or closed to match the
GridView hour, except non-conforming loads whose key is in the EPC (they take its status). Loads outside the main
island are also turned off (scale.turn_off_islanded()). The rules (load and gen targets per area, unit commitment,
make-up, balancing the swing) are in ReadMe.md, "Area Mode".

Working tables are in the shape Scripts/scale.py expects (see its docstring), plus AreaNumber, GenType and Method.
"""
import numpy as np
import pandas as pd
from . import prep, pw, scale

# ------------------ GenType ------------------
GEN_TYPES = ['Solar', 'Hydro', 'Wind', 'Thermal', 'Battery Storage', 'Unknown']

# Unit type letter code -> GenType. None: decided by the fuel type instead.
# From CaseBridge core/schemas/codes/generator_codes.json (technology list, its `group`), which also has each code's
# EPC and PowerWorld integer. Deviations from those groups are marked DECIDE: they are judgment calls for the owner.
UNIT_GENTYPE: dict[str, str | None] = {
    # Thermal
    'ST': 'Thermal', 'CA': 'Thermal', 'XC': 'Thermal', 'CC': 'Thermal', 'IC': 'Thermal', 'IT': 'Thermal',
    'GT': 'Thermal', 'JE': 'Thermal', 'CS': 'Thermal', 'BT': 'Thermal', 'CT': 'Thermal',
    # Hydro
    'HY': 'Hydro', 'HA': 'Hydro', 'HB': 'Hydro', 'HK': 'Hydro',
    'PS': 'Hydro',              # DECIDE: pumped storage. CaseBridge groups it as Storage.
    # Wind
    'WT': 'Wind', 'W1': 'Wind', 'W2': 'Wind', 'W3': 'Wind', 'W4': 'Wind', 'WS': 'Wind',
    # Solar
    'PV': 'Solar', 'P2': 'Solar', 'P3': 'Solar', 'P4': 'Solar',
    'CP': 'Solar',              # DECIDE: concentrated solar. CaseBridge groups it as Storage.
    # Battery Storage
    'BA': 'Battery Storage', 'FW': 'Battery Storage', 'ES': 'Battery Storage', 'CE': 'Battery Storage',
    # Not generation by type: never decided by fuel.
    'SC': 'Unknown', 'SV': 'Unknown', 'MP': 'Unknown',  # DECIDE: condensers, SVCs, motor/pumps.
    'DC': 'Unknown',            # DC ties: never scaled (MANUAL_UNIT_TYPES).
    # Unit type unstated or generic: use the fuel.
    'UN': None, 'OT': None, 'FC': None, '': None,
}

# Fuel type letter code -> GenType, used when the unit type gives None (or is not in UNIT_GENTYPE).
# From the same file (fuel list, its `group`): Coal, oil, gas, nuclear, geothermal, waste heat and biomass are Thermal.
FUEL_GENTYPE: dict[str, str] = {
    **{c: 'Thermal' for c in ['ANT', 'BIT', 'LIG', 'SGC', 'SUB', 'WC', 'RC',            # Coal
                              'DFO', 'JF', 'KER', 'PC', 'PG', 'RFO', 'WO',              # Oil
                              'SGP', 'BFG', 'NG', 'OG',                                 # Gas
                              'AB', 'MSW', 'OBS', 'WDS', 'OBL', 'SLW', 'BLQ', 'WDL', 'LFG', 'OBG',  # Biomass
                              'GEO', 'NUC', 'PUR', 'WH', 'TDF']},
    'WAT': 'Hydro', 'WPS': 'Hydro',
    'SUN': 'Solar',
    'WND': 'Wind',
    'MWH': 'Battery Storage',
    'UN': 'Unknown', 'OTH': 'Unknown', '': 'Unknown',
}

# Unit types left out of the scaling, in both cases: adjusted by hand.
MANUAL_UNIT_TYPES = {'DC'}

# Stop before scaling if more than this share of the EPC's generation (|MW|) has no GenType: its unit & fuel codes are
# probably blank, and every typed bucket in the PWB would be scaled to 0. 1.0 turns the check off.
MAX_UNKNOWN_SHARE = 0.5

def _code(s: pd.Series) -> pd.Series:
    """'HY', 'hy ', 'HY (Hydraulic Turbine)' -> 'HY'. Missing -> ''."""
    return s.fillna('').astype(str).str.strip().str.split(' ', n=1).str[0].str.upper()

def classify_gen_type(df: pd.DataFrame) -> pd.Series:
    """GenType of each gen, from its UnitTypeCode, else its FuelTypeCode. Codes in neither map are Unknown."""
    by_unit = _code(df['UnitTypeCode']).map(UNIT_GENTYPE)
    by_fuel = _code(df['FuelTypeCode']).map(FUEL_GENTYPE)
    return by_unit.fillna(by_fuel).fillna('Unknown')

# ------------------ Fields read from each case ------------------
_KEY = {'ObjectID': str, 'BusNum': int, 'BusName': str, 'NomkV': float, 'ID': str, 'AreaNumber': int}
LOAD_FIELDS: dict[str, type] = {**_KEY, 'Status': str, 'AGC': str, 'DistStatus': str, **{c: float for c in scale.SCALED['Load']}}
GEN_FIELDS: dict[str, type] = {**_KEY, 'Status': str, 'MW': float, 'MWSetPoint': float, 'MWMax': float, 'MWMin': float,
                               'UnitTypeCode': str, 'FuelTypeCode': str}

LOAD_P = ['SMW', 'IMW', 'ZMW']
LOAD_Q = ['SMvar', 'IMvar', 'ZMvar']

def read_case(SimAuto) -> dict[str, pd.DataFrame]:
    """Loads & gens of the open case (PWB, or an EPC opened in PowerWorld)."""
    return {'Load': pw.get_df(SimAuto, 'Load', LOAD_FIELDS), 'Gen': pw.get_df(SimAuto, 'Gen', GEN_FIELDS)}

def closed(df: pd.DataFrame) -> pd.Series:
    return df['Status'] == 'Closed'

# ------------------ Load targets ------------------
def sign_totals(value: pd.Series, area: pd.Series, name: str) -> pd.DataFrame:
    """Per area: total of the positive values ('<name> +') and of the negative values ('<name> -')."""
    return pd.DataFrame({f'{name} +': value.clip(lower=0).groupby(area).sum(),
                         f'{name} -': value.clip(upper=0).groupby(area).sum()})

def epc_only_negative(epc: pd.DataFrame, case_load: pd.DataFrame, col: str) -> pd.Series:
    """EPC loads with a negative `col` whose key isn't in the PWB (e.g. GridView's X1/X2 distributed generation loads)."""
    return ~epc['ObjectID'].isin(case_load['ObjectID']) & (epc[col] < 0)

def netted_sign_totals(epc: pd.DataFrame, case_load: pd.DataFrame, col: str, name: str) -> pd.DataFrame:
    """
    The EPC's per-area positive and negative totals of `col` (sign_totals()), with EPC-only negative loads netted into
    the positive total: the PWB has nothing to place them on, so they lower the area's positive load instead.
    """
    netted = epc_only_negative(epc, case_load, col)
    totals = sign_totals(epc[col].where(~netted, 0.0), epc['AreaNumber'], name)
    totals[f'{name} +'] += epc[col].where(netted, 0.0).groupby(epc['AreaNumber']).sum().reindex(totals.index, fill_value=0.0)
    return totals

def sign_share(value: pd.Series, area: pd.Series, target_pos: pd.Series, target_neg: pd.Series,
               fallback: pd.Series | None = None) -> pd.Series:
    """
    New value of each load. Per area, the positive loads are scaled by one factor so their total meets target_pos,
    and the negative loads by another so theirs meets target_neg (0 for an area missing from a target).
    A sign with no seed loads in an area can't be scaled, unless `fallback` (weights > 0) is given: then its target is
    spread over that area's zero-valued loads in proportion to fallback. Anything else is left unmet.
    """
    new = pd.Series(0.0, index=value.index)
    for sign, target in ((1, target_pos), (-1, target_neg)):
        part = (value * sign).clip(lower=0)
        t = (target.reindex(area.unique()).fillna(0.0) * sign).clip(lower=0)
        seed = part.groupby(area).sum()
        k = (t / seed).where(seed > 0, 0.0)
        new += sign * part * area.map(k)
        if fallback is not None:
            w = fallback.clip(lower=0).where(value == 0, 0.0)
            w_sum = w.groupby(area).sum()
            k_fb = (t / w_sum).where((seed <= 0) & (w_sum > 0), 0.0)
            new += sign * w * area.map(k_fb)
    return new

# Non-conforming loads: PowerWorld's load AGC field is NO (the RAW SCALE flag 0; an EPC's nonc flag 1, presumably).
# False: they are left out of both cases' area totals. Those whose key is in the EPC take its status and MW/Mvar 1:1
# (nonconforming_targets()); the others keep their seed values.
SCALE_NONCONFORMING = False

def conforming(load: pd.DataFrame) -> pd.Series:
    return load['AGC'].fillna('').astype(str).str.strip().str.upper() != 'NO'

def nonconforming_targets(epc_load: pd.DataFrame, case_load: pd.DataFrame) -> pd.DataFrame:
    """
    Target rows (load_targets()' shape, Method 'EPC 1:1') for non-conforming PWB loads whose key is in the EPC: its
    status, its S + I + Z MW and Mvar as constant power (0 if Open there), no current, impedance or dist gen. Any seed
    status: one Closed in the EPC is closed in by the engine (at 0 MW, then ramped), one Open there is ramped down and opened.
    """
    epc = epc_load.set_index('ObjectID')
    e = epc.loc[case_load['ObjectID']]
    on = (e['Status'] == 'Closed').to_numpy()
    df = case_load.copy()
    df['Method'] = 'EPC 1:1'
    df['Basis MW'] = np.where(on, e[LOAD_P].sum(axis=1), 0.0)
    df['Basis Mvar'] = np.where(on, e[LOAD_Q].sum(axis=1), 0.0)
    for c in ['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']:
        df[f'{c}_Target'] = 0.0
    df['SMW_Target'], df['SMvar_Target'] = df['Basis MW'], df['Basis Mvar']
    df['Planned MW'] = df['SMW_Target']
    df['Status_Target'] = e['Status'].to_numpy()
    df['DistStatus_Target'] = 'Open'
    for c in [*scale.STATUS['Load'], *scale.SCALED['Load']]:
        df[f'{c}_Seed'] = df[c]
    df['ExclusionReason'] = ''
    df['Nonconforming'] = True
    return df.set_index('ObjectID')

def load_targets(epc_load: pd.DataFrame, case_load: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Returns (targets, area report). targets holds the case's Closed loads. Open loads aren't scaled.
    """
    epc = epc_load.assign(P=epc_load[LOAD_P].sum(axis=1).where(closed(epc_load), 0.0),
                          Q=epc_load[LOAD_Q].sum(axis=1).where(closed(epc_load), 0.0))
    df = case_load[closed(case_load)].copy()
    df['P'] = df[LOAD_P].sum(axis=1)
    df['Q'] = df[LOAD_Q].sum(axis=1)

    # Basis: a load whose key is in the EPC starts from its EPC value (0 if Open there), any other from its seed value.
    epc_by_id = epc.set_index('ObjectID')
    in_epc = df['ObjectID'].isin(epc_by_id.index)
    df['Method'] = np.where(in_epc, 'Key', 'Seed')
    df['Basis MW'] = df['ObjectID'].map(epc_by_id['P']).where(in_epc, df['P'])
    df['Basis Mvar'] = df['ObjectID'].map(epc_by_id['Q']).where(in_epc, df['Q'])

    # Then each area's positive and negative basis is scaled to the EPC's totals, for MW and for Mvar. See sign_share().
    # Where every key matches, the factors are 1 and each load gets its own EPC value.
    tgt = pd.concat([netted_sign_totals(epc, case_load, 'P', 'MW'), netted_sign_totals(epc, case_load, 'Q', 'Mvar')], axis=1)
    seed = pd.concat([sign_totals(df['P'], df['AreaNumber'], 'MW'), sign_totals(df['Q'], df['AreaNumber'], 'Mvar')], axis=1)
    df['SMW_Target'] = sign_share(df['Basis MW'], df['AreaNumber'], tgt['MW +'], tgt['MW -'])
    # Mvar of a sign which no load in the area has is spread over its zero-Mvar loads in proportion to |MW|.
    df['SMvar_Target'] = sign_share(df['Basis Mvar'], df['AreaNumber'], tgt['Mvar +'], tgt['Mvar -'], fallback=df['Basis MW'].abs())

    for c in ['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']:
        df[f'{c}_Target'] = 0.0
    df['Planned MW'] = df['SMW_Target']             # What the plan placed. Never rescaled: areas() measures the balances against it.
    df['Status_Target'] = df['Status']
    df['DistStatus_Target'] = 'Open'
    for c in [*scale.STATUS['Load'], *scale.SCALED['Load']]:
        df[f'{c}_Seed'] = df[c]
    df['ExclusionReason'] = ''
    df['Nonconforming'] = False

    planned = pd.concat([sign_totals(df['SMW_Target'], df['AreaNumber'], 'MW'),
                         sign_totals(df['SMvar_Target'], df['AreaNumber'], 'Mvar')], axis=1)
    report = pd.DataFrame({
        'Target MW': tgt['MW +'] + tgt['MW -'], 'Target Mvar': tgt['Mvar +'] + tgt['Mvar -'],
        'Seed MW': seed['MW +'] + seed['MW -'], 'Seed Mvar': seed['Mvar +'] + seed['Mvar -'],
        'Seed Dist MW': df.groupby('AreaNumber')['DistMWInput'].sum(),
        'Planned MW': planned['MW +'] + planned['MW -'], 'Planned Mvar': planned['Mvar +'] + planned['Mvar -'],
        **{f'Target {c}': tgt[c] for c in tgt}, **{f'Seed {c}': seed[c] for c in seed},
        **{f'Planned {c}': planned[c] for c in planned},
        'Basis MW': df.groupby('AreaNumber')['Basis MW'].sum(),
        'Seed Loads': df.groupby('AreaNumber').size(),
        'Key Loads': in_epc.groupby(df['AreaNumber']).sum(),
        'EPC Loads': epc[closed(epc)].groupby('AreaNumber').size(),
        # e.g. GridView's X1/X2 distributed generation loads, which aren't in the PWB.
        'EPC-only MW': epc[~epc['ObjectID'].isin(case_load['ObjectID'])].groupby('AreaNumber')['P'].sum(),
        'EPC-only Netted MW': epc['P'].where(epc_only_negative(epc, case_load, 'P'), 0.0).groupby(epc['AreaNumber']).sum(),
    }).fillna(0.0)
    report['Unmet MW'] = report['Target MW'] - report['Planned MW']
    report['Unmet Mvar'] = report['Target Mvar'] - report['Planned Mvar']
    for c in tgt:
        report[f'Unmet {c}'] = report[f'Target {c}'] - report[f'Planned {c}']
    report.index.name = 'AreaNumber'
    return df.drop(columns=['P', 'Q']).set_index('ObjectID'), report.reset_index()

# ------------------ Gen targets ------------------
def capacity_share(target: float, cap: pd.Series) -> pd.Series:
    """Spreads target (>= 0) over units in proportion to cap (>= 0): the same fraction of cap on each, at most all of it."""
    total = cap.sum()
    return cap * (min(target / total, 1.0) if total > 0 else 0.0)

def gen_keys(df: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
    """(AreaNumber, GenType) of each gen; one area / one group for a table without the columns (some test tables)."""
    area = df['AreaNumber'] if 'AreaNumber' in df else pd.Series(0, index=df.index)
    group = classify_gen_type(df) if 'UnitTypeCode' in df and len(df) else pd.Series('All', index=df.index, dtype=object)
    return area, group

def proportional_fill(target: float, weight: np.ndarray, cap: np.ndarray) -> np.ndarray:
    """
    Spreads target (>= 0) over units in proportion to weight (>= 0), at most cap each (>= 0): units which reach their cap
    stay there and the rest is spread again over the others. Units with weight 0 get 0. Less than target if all are capped.
    """
    x = np.zeros(len(weight))
    free = weight > 0
    remaining = target
    while free.any() and remaining > 1e-9:
        share = np.where(free, remaining * weight / weight[free].sum(), 0.0)
        over = free & (share > cap)
        if not over.any():
            x[free] = share[free]
            break
        x[over] = cap[over]
        remaining -= cap[over].sum()
        free &= ~over
    return x

def spread_by_group(ask: pd.Series, mw: pd.Series, cap: pd.Series, area: pd.Series, group: pd.Series,
                    spill_all: bool = True) -> tuple[pd.Series, float]:
    """
    Places ask (MW, + or -, a Series by (AreaNumber, Group)) on gens with output mw (> 0), area and group, in proportion
    to mw, never above cap or below 0. What a group's own units can't take is spread over the rest of its area's
    gens, and (spill_all) what the area can't take over every gen. At the area stage, what is left going up and what
    is left going down are placed separately, each away from the groups asked to go the other way, so one group's
    move never undoes another's. Returns (MW change per gen, MW not placed: net of up and down).
    """
    m, c = mw.to_numpy(dtype=float), np.nan_to_num(cap.to_numpy(dtype=float), nan=0.0)
    x = np.zeros(len(m))
    by_group = pd.Series(m).groupby([area.to_numpy(), group.to_numpy()]).indices if len(m) else {}
    by_area = pd.Series(m).groupby(area.to_numpy()).indices if len(m) else {}
    def place(a, rows):
        if abs(a) < 1e-9 or not len(rows):
            return a
        if a > 0:
            add = proportional_fill(a, m[rows], np.clip(c[rows] - m[rows] - x[rows], 0.0, None))
        else:
            add = -proportional_fill(-a, m[rows], np.clip(m[rows] + x[rows], 0.0, None))
        x[rows] += add
        return a - add.sum()
    left: dict = {}         # area -> {sign: (MW left, groups asked that way)}
    for (a, g), v in ask.items():
        rest = place(v, by_group.get((a, g), []))
        sign = 1 if v > 0 else -1
        mw_left, groups = left.setdefault(a, {}).get(sign, (0.0, set()))
        left[a][sign] = (mw_left + rest, groups | {g})
    total = 0.0
    for a, signs in left.items():
        rows = np.asarray(by_area.get(a, []), dtype=int)
        for sign, (v, _) in signs.items():
            other = signs.get(-sign, (0.0, set()))[1]
            total += place(v, rows[~np.isin(group.to_numpy()[rows], list(other))] if len(rows) else rows)
    if spill_all:
        total = place(total, np.arange(len(m)))
    return pd.Series(x, index=mw.index), float(total)

def spread_evenly(total: float, mw: pd.Series, cap: pd.Series, area: pd.Series, group: pd.Series) -> tuple[pd.Series, float]:
    """total MW onto gens by one factor on mw (each area & group asked its share of total by its mw), spread_by_group()."""
    if not mw.sum() > 0:
        return pd.Series(0.0, index=mw.index), float(total)
    ask = (mw * (total / mw.sum())).groupby([area, group]).sum()
    return spread_by_group(ask, mw, cap, area, group)

def scalable_gens(df: pd.DataFrame) -> pd.DataFrame:
    """All gens but the giant swing and MANUAL_UNIT_TYPES."""
    return df[(df['ID'] != prep.SWING_ID) & ~_code(df['UnitTypeCode']).isin(MANUAL_UNIT_TYPES)]

# ------------------ Unit commitment ------------------
# Per (area, GenType) bucket, which PWB units are online is chosen to fit the target, rather than kept from the seed.
COMMIT = True       # False: keep every unit's seed status.
# Every unit online, all at the same % output (run of the river hydro, variable resources, storage).
ALL_ON_TYPES = {'Hydro', 'Wind', 'Solar', 'Battery Storage'}
# At a target of 0, these go offline. The other ALL_ON_TYPES keep their seed statuses, at 0 MW.
OFF_AT_ZERO_TYPES = {'Wind', 'Solar'}
# Other types (Thermal, Unknown) start from the seed's online units, so that the target is at most COMMIT_MAX_CF of the
# online MWMax and no more units are online than that needs. Units online in the EPC (same key) are preferred:
# scaling up brings on offline units which are online in the EPC first, then the biggest;
# scaling down takes off online units which are offline in the EPC first, then the smallest.
COMMIT_MAX_CF = 0.9
# Warn when a group is scaled below this fraction of its basis: its units stay online at low output (maybe under MWMin).
LOW_FACTOR_WARNING = 0.3

def commit(rows: pd.DataFrame, target: float, gen_type: str, epc_online: set) -> pd.Series:
    """
    Which units of one bucket are online (True) for `target` MW. Only units with capacity in the direction of the
    target count (MWMax > 0, or MWMin < 0 for a negative target); units without keep their seed status.
    """
    cap = (rows['MWMax'] if target >= 0 else -rows['MWMin']).fillna(0.0).clip(lower=0)
    able = cap > 0
    online = closed(rows).copy()
    if gen_type in ALL_ON_TYPES:
        if target != 0:
            online[able] = True
        elif gen_type in OFF_AT_ZERO_TYPES:
            online[able] = False
        return online
    if target == 0:     # Not by the loop below: the sum of what is left can be off by float residue.
        online[able] = False
        return online

    need = abs(target) / COMMIT_MAX_CF
    on = online[able]
    have = cap[on[on].index].sum()
    in_epc = rows['ObjectID'].isin(epc_online)
    if have < need:     # Scaling up: online in the EPC first, then the biggest.
        off = on[~on].index
        for i in pd.DataFrame({'epc': in_epc[off], 'cap': cap[off]}).sort_values(['epc', 'cap'], ascending=False).index:
            if have >= need:
                break
            online[i] = True
            have += cap[i]
    else:               # Scaling down: offline in the EPC first, then the smallest, while the rest still cover the target.
        up = on[on].index
        for i in pd.DataFrame({'epc': in_epc[up], 'cap': cap[up]}).sort_values(['epc', 'cap']).index:
            if have - cap[i] < need:
                continue    # Too big to take off: a smaller one later may still go.
            online[i] = False
            have -= cap[i]
    return online

def dispatch_bucket(rows: pd.DataFrame, target: float, gen_type: str, epc_online: set) -> tuple[pd.Series, pd.Series, str]:
    """
    (Status_Target, MWSetPoint_Target, method) of one (area, GenType) bucket's units, from their basis (gen_targets()).
    'Scaled': the online units' basis MW is scaled by one factor to the target, capped at MWMax (the rest spread over
    the others). Any MW still left brings offline units online by commit() and shares it over them by capacity.
    'Capacity': only when there is nothing to scale (target 0, or no online basis MW in the target's direction):
    units are committed by commit() and share the target in proportion to MWMax (|MWMin| for a negative target),
    following each GenType's dispatch rules.
    (Every group scales from its basis, keys or not: decided by Fable 2026-11-15. The seed dispatch is a solved operating
    point, scaling it stays close to that, and one rule for all groups is predictable.)
    Units with a basis of the opposite sign (e.g. storage charging in a generating bucket) keep it.
    """
    status = rows['Basis Status'].copy()
    basis = rows['Basis MW']
    sign = 1.0 if target >= 0 else -1.0
    cap = (rows['MWMax'] if target >= 0 else -rows['MWMin']).fillna(0.0).clip(lower=0)
    along = (basis * sign).clip(lower=0)                # Basis in the target's direction.
    against = basis.where(basis * sign < 0, 0.0)        # Kept as it is.
    want = abs(target - against.sum())

    if target != 0 and along.sum() > 0:
        factor = want / along.sum()
        if factor < LOW_FACTOR_WARNING:
            print(f'WARNING: area {rows["AreaNumber"].iloc[0]} {gen_type} scaled to {factor:.0%} of its basis; '
                  f'its units stay online at low output.')
        fill = proportional_fill(want, along.to_numpy(float), cap.where(cap > along, along).to_numpy(float))
        mw = against + sign * pd.Series(fill, index=rows.index)
        left = want - (mw - against).abs().sum()
        if left > 1e-6 and COMMIT:
            off = rows[status != 'Closed']
            extra = commit(off.assign(Status='Open'), sign * left, gen_type, epc_online)
            extra = extra[extra].index
            mw[extra] = sign * capacity_share(left, cap[extra])
            status[extra] = 'Closed'
        return status, mw, 'Scaled'

    if COMMIT:
        status = pd.Series(np.where(commit(rows.assign(Status=status), target, gen_type, epc_online), 'Closed', 'Open'),
                           index=rows.index)
    on = status == 'Closed'
    mw = pd.Series(0.0, index=rows.index)
    mw[on] = sign * capacity_share(abs(target), cap[on])
    return status, mw, 'Capacity'

def gen_targets(epc_gen: pd.DataFrame, case_gen: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Returns (targets, bucket report). targets holds the case's gens which are online now or after commitment,
    except the giant swing and DC ties.
    """
    epc_gen, case_gen = scalable_gens(epc_gen), scalable_gens(case_gen)
    epc = epc_gen.assign(GenType=classify_gen_type(epc_gen), P=epc_gen['MWSetPoint'].fillna(0.0).where(closed(epc_gen), 0.0))
    unknown, total = epc.loc[epc['GenType'] == 'Unknown', 'P'].abs().sum(), epc['P'].abs().sum()
    if unknown > MAX_UNKNOWN_SHARE * total:
        raise ValueError(f'{unknown:.0f} MW of the EPC\'s {total:.0f} MW of generation has no GenType (its unit and fuel codes '
                         f'are blank or unknown). Scaling would take the rest of the fleet offline. Map the codes in '
                         f'[area.unit_gentype] / [area.fuel_gentype] in settings.toml, or raise [area] max_unknown_share.')
    epc_online = set(epc.loc[closed(epc), 'ObjectID'])

    df = case_gen.copy()
    df['GenType'] = classify_gen_type(df)
    # Basis: a unit whose key is in the EPC starts from its EPC status (with COMMIT) and MW, any other from its seed's.
    by_id = epc.set_index('ObjectID')
    in_epc = df['ObjectID'].isin(by_id.index)
    df['Method'] = np.where(in_epc, 'Key', 'Seed')
    df['Basis Status'] = df['ObjectID'].map(by_id['Status']).where(in_epc & COMMIT, df['Status'])
    df['Basis MW'] = df['ObjectID'].map(by_id['P']).where(in_epc, df['MWSetPoint']).where(df['Basis Status'] == 'Closed', 0.0)
    df['Status_Target'] = df['Basis Status']
    df['MWSetPoint_Target'] = df['Basis MW']

    keys = ['AreaNumber', 'GenType']
    tgt = epc.groupby(keys)['P'].sum()
    bucket_method = {}
    for bucket, rows in df.groupby(keys):
        t = float(tgt.get(bucket, 0.0))
        status, mw, bucket_method[bucket] = dispatch_bucket(rows, t, bucket[1], epc_online)
        df.loc[rows.index, 'Status_Target'] = status
        df.loc[rows.index, 'MWSetPoint_Target'] = mw.where(status == 'Closed', 0.0)
    own = df.groupby(keys)['MWSetPoint_Target'].sum()       # What each bucket's own units took.
    own_mw = df['MWSetPoint_Target'].copy()
    if MAKEUP:
        # What a bucket's own units can't take (no PWB units, or capped at MWMax) goes onto the area's other GenTypes:
        # first their online units, in proportion to MW; then, for what is still left going up, the other buckets with
        # offline units are re-dispatched with a bigger target, so commit() brings units online by the usual rules.
        unmet = tgt.sub(own, fill_value=0.0)
        unmet = unmet[unmet.abs() >= 1e-6]
        on = df[(df['Status_Target'] == 'Closed') & (df['MWSetPoint_Target'] > 0)]
        if len(unmet) and len(on):
            mw = on['MWSetPoint_Target']
            df.loc[on.index, 'MWSetPoint_Target'] += spread_by_group(
                unmet, mw, on['MWMax'].where(on['MWMax'] > mw, mw), on['AreaNumber'], on['GenType'], spill_all=False)[0]
        left = unmet.groupby(level='AreaNumber').sum().sub(
            (df['MWSetPoint_Target'] - own_mw).groupby(df['AreaNumber']).sum(), fill_value=0.0)
        for a, mw_left in left[left > 1e-6].items():
            short = set(unmet[unmet > 0].xs(a, level='AreaNumber').index) if a in unmet.index.get_level_values(0) else set()
            rows_a = df[(df['AreaNumber'] == a) & ~df['GenType'].isin(short)]
            off_cap = rows_a['MWMax'].where(rows_a['Status_Target'] != 'Closed', 0.0).clip(lower=0).groupby(rows_a['GenType']).sum()
            off_cap = off_cap[(off_cap > 0) & ~off_cap.index.isin(ALL_ON_TYPES)]    # Only dispatchable types: commit() would bring every Wind/Solar/Hydro/Storage unit online.
            for g, c in off_cap.items():
                rows = rows_a[rows_a['GenType'] == g]
                more = float(df.loc[rows.index, 'MWSetPoint_Target'].sum()) + mw_left * c / off_cap.sum()
                status, mw, _ = dispatch_bucket(rows, more, g, epc_online)
                df.loc[rows.index, 'Status_Target'] = status
                df.loc[rows.index, 'MWSetPoint_Target'] = mw.where(status == 'Closed', 0.0)
    df['Plan Make-up MW'] = df['MWSetPoint_Target'] - own_mw
    df['Planned MW'] = df['MWSetPoint_Target']      # What the plan placed. Never rescaled: areas() measures the balances against it.

    df['MWSetPoint'] = df['MWSetPoint'].where(closed(df), 0.0)   # Open units are at 0 MW, as the engine expects.
    df['Status_Seed'] = df['Status']
    df['MWSetPoint_Seed'] = df['MWSetPoint']
    df['ExclusionReason'] = ''

    on_seed, on_plan = closed(df), df['Status_Target'] == 'Closed'
    g = df.groupby(keys)
    report = pd.DataFrame({
        'Target MW': tgt,
        'Seed MW': g['MWSetPoint'].sum(),
        'Planned MW': g['MWSetPoint_Target'].sum(),
        'Make-up MW': g['Plan Make-up MW'].sum(),
        'Seed Online': on_seed.groupby([df[k] for k in keys]).sum(),
        'Planned Online': on_plan.groupby([df[k] for k in keys]).sum(),
        'Committed': (on_plan & ~on_seed).groupby([df[k] for k in keys]).sum(),
        'Decommitted': (on_seed & ~on_plan).groupby([df[k] for k in keys]).sum(),
        'Seed Online MWMax': df['MWMax'].where(on_seed, 0.0).groupby([df[k] for k in keys]).sum(),
        'Planned Online MWMax': df['MWMax'].where(on_plan, 0.0).groupby([df[k] for k in keys]).sum(),
        'All Units MWMax': g['MWMax'].sum(),
        'Basis MW': g['Basis MW'].sum(),
        'Key Gens': in_epc.groupby([df[k] for k in keys]).sum(),
        'EPC Online': epc[closed(epc)].groupby(keys).size(),
        'EPC Online MWMax': epc[closed(epc)].groupby(keys)['MWMax'].sum(),
    }).fillna(0.0)
    report = report[(report.drop(columns='All Units MWMax') != 0).any(axis=1)]   # Drop buckets with nothing online anywhere.
    report['Method'] = report.index.map(lambda k: bucket_method.get(k, ''))
    # Unmet: what the bucket's own units couldn't take. Planned includes what it took for the area's other buckets.
    report['Unmet MW'] = report['Target MW'] - (report['Planned MW'] - report['Make-up MW'])
    df = df.drop(columns='Plan Make-up MW')
    # Units offline before and after aren't touched.
    df = df[on_seed | on_plan]
    return df.set_index('ObjectID'), report.reset_index()

def build_targets(epc: dict[str, pd.DataFrame], case: dict[str, pd.DataFrame], swing_buses=(),
                  island_by_bus: pd.Series | None = None) -> tuple[dict, dict]:
    """
    Returns (targets {'Gen', 'Load'} for scale.iterate_to_gen_load_targets(), reports {'area_load', 'area_gen', ...}).
    reports['nonconforming']: ObjectIDs of the main island's non-conforming loads (flagged in either case; empty with
    SCALE_NONCONFORMING). They are out of the area totals, the balances and the make-up; those with a key in the EPC are
    target rows (Nonconforming True, Method 'EPC 1:1': the EPC's status and MW/Mvar), the rest are left as is.
    Gens on swing_buses (the slack bus) are left out of both cases: the solution sets their MW, and swing balancing
    deals with what they carry.
    island_by_bus (scale.seed_islands()): the PWB's gens/loads outside the main island take no share of the targets
    (the main island's units carry the whole area target), and are left out of the area seed totals, like
    non-conforming loads. Those Closed in the seed are added back as excluded rows targeted Open at 0 MW (turned off
    before scaling), so the log names them (scale.island_reason()); reports['islanded'] has their seed MW per
    (AreaNumber, Group).
    """
    case_reason = {t: scale.island_reason(case[t], island_by_bus) for t in ('Gen', 'Load')}
    islanded = {t: case[t][case_reason[t] != ''].assign(ExclusionReason=case_reason[t]) for t in ('Gen', 'Load')}
    case = {t: case[t][case_reason[t] == ''] for t in ('Gen', 'Load')}
    epc_load, case_load = epc['Load'], case['Load']
    # Non-conforming is decided by key: a PWB load flagged in either case is out of both cases' totals (and so is an
    # EPC load flagged, or with a flagged PWB key).
    if SCALE_NONCONFORMING:
        nonc_ids, epc_nonc = set(), pd.Series(False, index=epc_load.index)
    else:
        epc_flagged = ~conforming(epc_load)
        flagged_in_either = ~conforming(case_load) | case_load['ObjectID'].isin(epc_load.loc[epc_flagged, 'ObjectID'])
        nonc_ids = set(case_load.loc[flagged_in_either, 'ObjectID'])
        epc_nonc = epc_flagged | epc_load['ObjectID'].isin(nonc_ids)
    is_nonc = case_load['ObjectID'].isin(nonc_ids)
    case_nonc = case_load[is_nonc]
    load, load_report = load_targets(epc_load[~epc_nonc], case_load[~is_nonc])
    # Those with a key in the EPC take its status and MW 1:1 (target rows), unless Open in both cases (nothing to
    # change); the rest keep their seed values.
    epc_status = case_nonc['ObjectID'].map(epc_load.set_index('ObjectID')['Status'])
    keyed = epc_status.notna() & ~((epc_status == 'Open') & (case_nonc['Status'] == 'Open'))
    nonc_rows = nonconforming_targets(epc_load, case_nonc[keyed])
    load = pd.concat([load, nonc_rows])
    load_report = _nonconforming_report(load_report, case_nonc, keyed, nonc_rows)
    off_swing = lambda df: df[~df['BusNum'].isin(swing_buses)]
    gen, gen_report = gen_targets(off_swing(epc['Gen']), off_swing(case['Gen']))
    gen, load, islanded_mw = _add_islanded(gen, load, islanded['Gen'], islanded['Load'])
    return {'Gen': gen, 'Load': load}, {'area_load': load_report, 'area_gen': gen_report, 'islanded': islanded_mw,
                                        'nonconforming': nonc_ids, 'swing': swing_epc(epc['Gen'], case['Gen'], swing_buses)}

def _nonconforming_report(load_report: pd.DataFrame, case_nonc: pd.DataFrame, keyed: pd.Series,
                          nonc_rows: pd.DataFrame) -> pd.DataFrame:
    """
    load_report with the non-conforming columns (MW, per area): 'Left As Is MW' (seed MW of those which are not target
    rows), 'Set To EPC MW' and 'Set To EPC Loads' (their 1:1 target rows). An area with only non-conforming loads gets a row too.
    """
    mw = lambda df: df[closed(df)].assign(P=lambda d: d[LOAD_P].sum(axis=1)).groupby('AreaNumber')['P'].sum()
    cols = pd.DataFrame({'Left As Is MW': mw(case_nonc[~keyed]),
                         'Set To EPC MW': nonc_rows.groupby('AreaNumber')['SMW_Target'].sum(),
                         'Set To EPC Loads': nonc_rows.groupby('AreaNumber').size()})
    report = load_report.set_index('AreaNumber')
    index = report.index.union(pd.Index(case_nonc['AreaNumber'].unique(), name='AreaNumber'))
    report = report.reindex(index).fillna(0.0).join(cols.reindex(index)).fillna(0.0)
    return report.reset_index()

def swing_epc(epc_gen: pd.DataFrame, case_gen: pd.DataFrame, swing_buses) -> pd.DataFrame:
    """
    The case's own slack gens (Closed, on swing_buses) whose key is also in the EPC: 'Seed MW' (their solved MW, or
    MWSetPoint if MW isn't read) and 'EPC MW' (0 if Open in the EPC; kept within its MWMin to MWMax), by ObjectID. The balances aim each at its EPC MW
    (swing_goal()). Empty if the case has a giant swing: that one always aims at 0.
    """
    empty = pd.DataFrame(columns=['Seed MW', 'EPC MW'], dtype=float)
    if (case_gen['ID'] == prep.SWING_ID).any():
        return empty
    slack = case_gen[closed(case_gen) & case_gen['BusNum'].isin(swing_buses) & case_gen['ObjectID'].isin(epc_gen['ObjectID'])]
    if not len(slack):
        return empty
    epc = epc_gen.set_index('ObjectID')
    epc_mw = epc['MWSetPoint'].fillna(0.0).where(closed(epc), 0.0)
    seed = slack['MW'] if 'MW' in slack else slack['MWSetPoint']
    lo, hi = _limits(slack)
    goal = slack['ObjectID'].map(epc_mw).clip(lower=lo, upper=hi)     # Within its own limits (swing_goal()).
    return pd.DataFrame({'Seed MW': seed.to_numpy(dtype=float), 'EPC MW': goal.to_numpy(dtype=float)},
                        index=pd.Index(slack['ObjectID'], name='ObjectID'))

def _add_islanded(gen: pd.DataFrame, load: pd.DataFrame, i_gen: pd.DataFrame, i_load: pd.DataFrame):
    """
    build_targets(): the islanded gens/loads Closed in the seed (a second island's slack gens too; not the giant
    swing), as excluded target rows targeted Open at 0 MW, which step 02 writes before scaling
    (scale.turn_off_islanded()). Returns (gen, load, their seed MW per (AreaNumber, Group)).
    """
    i_gen = i_gen[i_gen['ID'] != prep.SWING_ID]
    i_gen, i_load = i_gen[closed(i_gen)].copy(), i_load[closed(i_load)].copy()
    for df, table in ((i_gen, 'Gen'), (i_load, 'Load')):
        for c in [*scale.STATUS[table], *scale.SCALED[table]]:
            df[f'{c}_Seed'] = df[c]
            df[f'{c}_Target'] = df[c]
        scale.turn_off_targets(df, table, df.index == df.index)
    i_gen['GenType'] = classify_gen_type(i_gen)
    i_gen['Planned MW'], i_load['Planned MW'] = 0.0, 0.0
    i_load['Nonconforming'] = False
    mw = pd.concat([_by_group(i_gen, i_gen['MWSetPoint'], i_gen['GenType']), _by_group(i_load, i_load[LOAD_P].sum(axis=1))])
    scale.print_islanded({'Gen': len(i_gen), 'Load': len(i_load)}, set(i_gen['ExclusionReason']) | set(i_load['ExclusionReason']))
    join = lambda a, b: pd.concat([a, b.set_index('ObjectID')]) if len(b) else a
    return join(gen, i_gen), join(load, i_load), mw

def dist_gen_cleanup(case_load: pd.DataFrame) -> pd.DataFrame:
    """
    Every load with distributed generation still on: set it Open & 0. GridView ignores distributed generation, so no
    hour keeps any. Step 02 writes it after the main scaling, before the final balance (which takes up what it removes):
    by then it is left only on loads the scaling never touched (Open, non-conforming without an EPC key) or didn't
    finish (held, excluded).
    dist_gen_removed() is what that takes off the case.
    """
    on = (case_load['DistStatus'] == 'Closed') | (case_load['DistMWInput'] != 0) | (case_load['DistMvarInput'] != 0)
    df = case_load[on]
    return df.assign(DistStatus='Open', DistMWInput=0.0, DistMvarInput=0.0, DistMW_Seed=df['DistMWInput'].where(
        closed(df) & (df['DistStatus'] == 'Closed'), 0.0))

def dist_gen_removed(cleanup: pd.DataFrame) -> float:
    """MW of distributed generation a dist_gen_cleanup() table takes off the case (it injected only where Closed)."""
    return float(cleanup['DistMW_Seed'].sum()) if len(cleanup) else 0.0

AREAS_COLUMNS = ['AreaNumber', 'Group', 'Seed MW', 'Target MW', 'Unmet MW', 'Not Reached MW', 'Make-up MW', 'Balance MW', 'Final MW', 'Error MW',
                 'Target Mvar', 'Final Mvar', 'Online Seed', 'Online Final', 'Seed Loss MW', 'Est. Loss MW', 'Final Loss MW', 'Note']
LEDGER_TOL_MW = 1.0     # A row's Final further than this from its ledger is noted as unexplained.

def _by_group(df: pd.DataFrame, value, group=None) -> pd.Series:
    """Sum of `value` (aligned to df's rows) by (AreaNumber, Group). Group: GenType of a gen table, else 'Load'."""
    if group is None:
        group = pd.Series('Load', index=df.index)
    frame = pd.DataFrame({'AreaNumber': df['AreaNumber'].to_numpy(), 'Group': np.asarray(group, dtype=object),
                          'v': np.asarray(value, dtype=float)})
    if frame.empty:
        return pd.Series(dtype=float, index=pd.MultiIndex.from_arrays([[], []], names=['AreaNumber', 'Group']))
    return frame.groupby(['AreaNumber', 'Group'])['v'].sum()

def areas(final: dict[str, pd.DataFrame] | None, reports: dict[str, pd.DataFrame], targets: dict[str, pd.DataFrame] | None = None,
          balance_tables=(), swing_buses=(), final_loss: pd.Series | None = None, makeup_tables=()) -> pd.DataFrame:
    """
    The `areas` sheet: a row per (AreaNumber, Group), Group a GenType or 'Load' (AREAS_COLUMNS). The ledger per row:
        Final = Target - Unmet - Not Reached + Make-up + Balance
    Make-up: what the plan moved onto the group to cover another GenType of its area which its own units couldn't
    (reports['area_gen'] 'Make-up MW'), plus what the make-up run moved onto it (makeup_tables: makeup_targets() tables,
    after scaling), to cover what excluded units didn't reach.
    Not Reached: over the main scaling's rows, end-of-run target less end-of-run value (excluded, diverged, held units).
    Balance: what swing balancing moved onto the group = the midrun part (end-of-run `_Target` less Planned MW, main rows:
    midrun balances rescale the main targets) + the final part (value less seed in each final balance's table).
    final: read_case() of the scaled case, or None before scaling (Final columns blank). targets: the main tables after
    scaling (None: nothing reached or balanced). balance_tables: balance_targets() tables of the final balance rounds,
    after scaling (balance_swing(..., tables=)). Gens on swing_buses are left out, as in build_targets(). The non-conforming
    loads (reports['nonconforming']) are left out of every column but the Load row's Note.
    Losses go on each area's Load row: Seed and Est. from reports['area_loss'] (loss_estimate()), Final from final_loss
    (area_losses() of the scaled case). An area with losses but no load gets a Load row for them.
    """
    key = ['AreaNumber', 'Group']
    nonc = reports['nonconforming']
    plan = pd.concat([reports['area_gen'].rename(columns={'GenType': 'Group'}),
                      reports['area_load'].rename(columns={'Seed Loads': 'Seed Online'}).assign(Group='Load')]).set_index(key)
    parts: dict[str, pd.Series] = {}

    p_of = lambda df, cols: df[cols].sum(axis=1)
    if targets is not None:
        g, l = targets['Gen'], targets['Load']
        # Islanded rows are out of the ledger, as they are out of Seed, Target and Final (even one that stayed on).
        g, l = g[~scale.is_island_reason(g['ExclusionReason'])], l[~scale.is_island_reason(l['ExclusionReason']) & ~l.index.isin(nonc)]
        gg = classify_gen_type(g)
        not_reached = (_by_group(g, g['MWSetPoint_Target'] - g['MWSetPoint'], gg)
                       .add(_by_group(l, p_of(l, [f'{c}_Target' for c in LOAD_P]) - p_of(l, LOAD_P)), fill_value=0.0))
        end_target = (_by_group(g, g['MWSetPoint_Target'], gg).add(_by_group(l, p_of(l, [f'{c}_Target' for c in LOAD_P])), fill_value=0.0))
        planned = _by_group(g, g['Planned MW'], gg).add(_by_group(l, l['Planned MW']), fill_value=0.0)
        parts['Not Reached MW'] = not_reached
        parts['Balance MW'] = end_target.sub(planned, fill_value=0.0)        # The midrun part.
    for col, tables in (('Balance MW', balance_tables), ('Make-up MW', makeup_tables)):
        for t in tables:
            g, l = t['Gen'], t['Load']
            g = g[~g['BusNum'].isin(swing_buses)]       # Left out of Final MW too (a second slack bus beside the giant swing).
            start = parts[col] if col in parts else _by_group(g.iloc[:0], [])
            parts[col] = (start.add(_by_group(g, g['MWSetPoint'] - g['MWSetPoint_Seed'], classify_gen_type(g)), fill_value=0.0)
                          .add(_by_group(l, p_of(l, LOAD_P) - l['P_Seed']), fill_value=0.0))
    if final is not None:
        # Islanded units (excluded rows of the targets) are left out of Final MW, as they are out of Seed and Target.
        out_ids = set().union(*[df.index[scale.is_island_reason(df['ExclusionReason'])] for df in (targets or {}).values()])
        load = final['Load'][closed(final['Load']) & ~final['Load']['ObjectID'].isin(out_ids | nonc)]
        gen = scalable_gens(final['Gen'])
        # STATCOMs the scaling added (ID 'SC', not in the targets) are not generation; a seed gen with that ID is.
        added = (gen['ID'] == scale.STATCOM_ID) & ~gen['ObjectID'].isin(targets['Gen'].index if targets is not None else [])
        gen = gen[closed(gen) & ~gen['BusNum'].isin(swing_buses) & ~gen['ObjectID'].isin(out_ids) & ~added]
        parts['Final MW'] = _by_group(gen, gen['MWSetPoint'], classify_gen_type(gen)).add(_by_group(load, p_of(load, LOAD_P)), fill_value=0.0)
        parts['Final Mvar'] = _by_group(load, p_of(load, LOAD_Q))
        parts['Online Final'] = _by_group(gen, 1.0, classify_gen_type(gen)).add(_by_group(load, 1.0), fill_value=0.0)

    index = plan.index
    for s in parts.values():
        index = index.union(s.index)
    loss = reports.get('area_loss', pd.DataFrame(columns=['Seed Loss MW', 'Est. Loss MW'], dtype=float))
    if final_loss is not None:
        loss = loss.join(final_loss.rename('Final Loss MW'), how='outer')
    loss = loss[(loss.fillna(0.0) != 0).any(axis=1)]
    if len(loss):
        index = index.union(pd.MultiIndex.from_arrays([loss.index, ['Load'] * len(loss)], names=key))
    index = index.sort_values()
    out = pd.DataFrame(index=index)
    for c in ('Seed MW', 'Target MW', 'Unmet MW', 'Target Mvar', 'Seed Online'):
        out[c] = plan[c].reindex(index) if c in plan else np.nan
    out[['Seed MW', 'Target MW', 'Unmet MW']] = out[['Seed MW', 'Target MW', 'Unmet MW']].fillna(0.0)
    out.loc[out.index.get_level_values('Group') != 'Load', 'Target Mvar'] = np.nan
    for c in ('Not Reached MW', 'Make-up MW', 'Balance MW'):
        out[c] = parts[c].reindex(index, fill_value=0.0) if c in parts else 0.0
    if 'Make-up MW' in plan:        # The plan's own area make-up (gen_targets()).
        out['Make-up MW'] += plan['Make-up MW'].reindex(index).fillna(0.0)
    out = out.rename(columns={'Seed Online': 'Online Seed'})
    out['Online Seed'] = out['Online Seed'].fillna(0.0)
    for c in ('Final MW', 'Final Mvar', 'Online Final'):
        out[c] = parts[c].reindex(index, fill_value=0.0) if final is not None else np.nan
    out.loc[out.index.get_level_values('Group') != 'Load', 'Final Mvar'] = np.nan
    out['Error MW'] = out['Final MW'] - out['Target MW']
    is_load = out.index.get_level_values('Group') == 'Load'
    areas_of = out.index.get_level_values('AreaNumber')
    for c in ('Seed Loss MW', 'Est. Loss MW', 'Final Loss MW'):
        col = loss[c] if c in loss else pd.Series(dtype=float)
        out[c] = np.where(is_load, col.reindex(areas_of).to_numpy() if c in loss else np.nan, np.nan)
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Make-up MW'] + out['Balance MW']
    out['Note'] = _area_notes(out, plan.reindex(index), ledger)
    islanded = reports.get('islanded', pd.Series(dtype=float)).reindex(index).fillna(0.0)
    still_on = pd.Series(0.0, index=index)
    if final is not None and targets is not None:     # Island units whose turn-off diverged: still Closed at the end.
        on_g = final['Gen'][closed(final['Gen']) & final['Gen']['ObjectID'].isin(out_ids)]
        on_l = final['Load'][closed(final['Load']) & final['Load']['ObjectID'].isin(out_ids)]
        still_on = (_by_group(on_g, on_g['MWSetPoint'], classify_gen_type(on_g)).add(_by_group(on_l, p_of(on_l, LOAD_P)), fill_value=0.0)
                    .reindex(index).fillna(0.0))
    note = pd.Series([('outside main island, turned off: ' + (f'{mw - on:.1f} MW' if not on else f'{mw:.1f} MW, of which {on:.1f} MW stayed on'))
                      if mw else '' for mw, on in zip(islanded, still_on)], index=index)
    out['Note'] = ['; '.join(filter(None, [n, i])) for n, i in zip(out['Note'], note)]
    return out.reset_index()[AREAS_COLUMNS]

def _area_notes(out: pd.DataFrame, plan: pd.DataFrame, ledger: pd.Series) -> list[str]:
    notes = []
    for i, k in enumerate(out.index):
        o, p, n = out.iloc[i], plan.iloc[i], []
        target, unmet = o['Target MW'], o['Unmet MW']
        if k[1] == 'Load':
            if target != 0 and p.get('Seed Online', 0) == 0:
                n.append('no PWB units')
            n_set, left = p.get('Set To EPC Loads', 0), p.get('Left As Is MW', 0)
            if pd.notna(n_set) and n_set:
                n.append(f"non-conforming set to EPC: {p['Set To EPC MW']:.1f} MW ({int(n_set)} load{'s' if n_set != 1 else ''})")
            if pd.notna(left) and left != 0:
                n.append(f'non-conforming left as is: {left:.1f} MW')
        else:
            units = p.get('All Units MWMax', 0) or p.get('Seed Online', 0)
            if target != 0 and not units:
                n.append('no PWB units')
            elif target == 0 and o['Seed MW'] != 0:
                n.append('no EPC target')
            elif abs(unmet) > 1e-6:
                n.append('capped at MWMax')
            if p.get('Method') == 'Capacity' and target != 0 and units:
                n.append('committed by capacity rules')
            basis = p.get('Basis MW', 0)
            if p.get('Method') == 'Scaled' and basis and 0 < (target - unmet) / basis < LOW_FACTOR_WARNING:
                n.append(f'scaled to {(target - unmet) / basis:.0%} of basis')
        if pd.notna(o['Final MW']) and abs(o['Final MW'] - ledger.iloc[i]) > LEDGER_TOL_MW:
            n.append(f"unexplained {o['Final MW'] - ledger.iloc[i]:.1f} MW")
        notes.append('; '.join(n))
    return notes

def compute_targets(SimAuto, gv_fp, pw_fp) -> tuple[dict, dict]:
    """Opens the EPC and the PWB, and returns build_targets() of their loads & gens. Closes the case."""
    pw.open_case(SimAuto, gv_fp)
    epc = read_case(SimAuto)
    pw.open_case(SimAuto, pw_fp)
    case = read_case(SimAuto)
    slack = prep.slack_buses(SimAuto)
    seed_loss = area_losses(SimAuto)
    islands = scale.seed_islands(SimAuto, slack)
    SimAuto.CloseCase()
    targets, reports = build_targets(epc, case, swing_buses=slack, island_by_bus=islands)
    reports['area_loss'] = loss_estimate(targets, reports['area_load'], seed_loss)
    return targets, reports

# ------------------ Losses ------------------
# Each area's losses are expected to go as (its target load / its seed load) ** LOSS_EXPONENT: I^2 R gives 2.
# 0: losses are taken to stay at their seed values. The areas sheet has Seed, Est. and Final Loss MW to check it by.
LOSS_EXPONENT = 2.0
LOSS_RATIO_MAX = 2.0    # An area's load ratio is capped at this (a small-load area could otherwise get huge losses).

def area_losses(SimAuto) -> pd.Series:
    """LossMW of each area of the open case, by AreaNumber."""
    df = pw.get_df(SimAuto, 'Area', {'Number': int, 'LossMW': float})
    return pd.Series(df['LossMW'].to_numpy(), index=pd.Index(df['Number'].to_numpy(), name='AreaNumber'), dtype=float)

def loss_estimate(targets: dict[str, pd.DataFrame], area_load: pd.DataFrame, seed_loss: pd.Series) -> pd.DataFrame:
    """
    Seed and estimated target LossMW per area (columns 'Seed Loss MW', 'Est. Loss MW', index AreaNumber). An area's load
    is its positive Closed load MW in the targets (negative loads, e.g. DC converters, left out; non-conforming loads
    with an EPC key are target rows, at their EPC values), plus the seed MW of the non-conforming loads left as is
    (area_load's 'Left As Is MW').
    The ratio is capped at LOSS_RATIO_MAX (printed). An area with no seed load keeps its seed losses.
    """
    l = targets['Load']
    l = l[scale.included(l)]
    def total(suffix):
        p = l[[c + suffix for c in LOAD_P]].sum(axis=1).where(l['Status' + suffix] == 'Closed', 0.0).clip(lower=0.0)
        return p.groupby(l['AreaNumber']).sum()
    fixed = area_load.set_index('AreaNumber')['Left As Is MW'].clip(lower=0.0)
    seed, target = total('').add(fixed, fill_value=0.0), total('_Target').add(fixed, fill_value=0.0)
    ratio = (target / seed).where(seed > 0, 1.0).reindex(seed_loss.index, fill_value=1.0)
    capped = ratio[ratio > LOSS_RATIO_MAX]
    if len(capped):
        print(f'Losses: load ratio capped at {LOSS_RATIO_MAX:g} in area(s) '
              + ', '.join(f'{a} ({r:.1f})' for a, r in capped.items()) + '.')
    ratio = ratio.clip(lower=0.0, upper=LOSS_RATIO_MAX)
    return pd.DataFrame({'Seed Loss MW': seed_loss, 'Est. Loss MW': seed_loss * ratio ** LOSS_EXPONENT})

# ------------------ Swing balancing ------------------
BALANCE_TOL_MW = 1.0        # Stop balancing once the swing is within this of 0 MW...
BALANCE_MAX_ROUNDS = 10     # ...or after this many rounds (losses move with each round, so it takes a few).
# While scaling (midrun_balancer()), balance once the swing goes outside its band:
BALANCE_TRIGGER_MW = 1000.0         # further than this from 0 MW (0: not used)...
BALANCE_AT_SWING_LIMITS = False     # ...or past the swing unit's own MWMax / MWMin (summed, if several).
MIN_OWN_RANGE_MW = 100.0            # With no trigger, the case's own slack needs a range at least this wide to be balanced midrun.


def swing_units(gen: pd.DataFrame, slack_buses) -> pd.DataFrame:
    """The giant swing (ID prep.SWING_ID) if the case has one Closed, else every Closed gen on a slack bus."""
    on = gen[closed(gen)]
    swing = on[on['ID'] == prep.SWING_ID]
    return swing if len(swing) else on[on['BusNum'].isin(slack_buses)]

def swing_mw(gen: pd.DataFrame, slack_buses) -> tuple[float, list]:
    """MW on the swing, and the swing's bus numbers."""
    swing = swing_units(gen, slack_buses)
    return float(swing['MW'].sum()), sorted(set(swing['BusNum']))

def _limits(units: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
    """(low, high) MW limits of each unit: MWMin / MWMax in order, a blank one unlimited."""
    a, b = units['MWMin'].astype(float).fillna(-np.inf), units['MWMax'].astype(float).fillna(np.inf)
    return np.minimum(a, b), np.maximum(a, b)

def swing_band(gen: pd.DataFrame, slack_buses) -> tuple[float, float]:
    """
    (low, high) MW the swing may sit within while scaling, before midrun_balancer() rebalances. The case's own slack
    gens (no giant swing): their MWMin to MWMax (summed). The giant swing: the narrowest of +/- BALANCE_TRIGGER_MW
    (if not 0) and its MWMin / MWMax (if BALANCE_AT_SWING_LIMITS).
    """
    swing = swing_units(gen, slack_buses)
    if len(swing) and not (swing['ID'] == prep.SWING_ID).any():
        lo, hi = _limits(swing)
        return float(lo.sum()), float(hi.sum())
    lo, hi = -np.inf, np.inf
    if BALANCE_TRIGGER_MW:
        lo, hi = -BALANCE_TRIGGER_MW, BALANCE_TRIGGER_MW
    if BALANCE_AT_SWING_LIMITS:
        swing = swing_units(gen, slack_buses)
        lo, hi = max(lo, float(swing['MWMin'].sum())), min(hi, float(swing['MWMax'].sum()))
    return lo, hi

def swing_goal(gen: pd.DataFrame, slack_buses, epc: pd.DataFrame | None = None) -> tuple[float, float]:
    """
    (low, high) MW the swing should end within. The giant swing: 0. The case's own slack gens: each one in `epc`
    (swing_epc()) at its EPC MW kept within its MWMin to MWMax (an EPC MW past a limit, or 0 for a slack Open in the
    EPC, aims at that limit), any other anywhere within its limits (a blank limit: unlimited); summed.
    """
    swing = swing_units(gen, slack_buses)
    if not len(swing) or (swing['ID'] == prep.SWING_ID).any():
        return 0.0, 0.0
    lo, hi = _limits(swing)
    epc_mw = swing['ObjectID'].map(epc['EPC MW']) if epc is not None and len(epc) else pd.Series(np.nan, index=swing.index)
    matched = epc_mw.notna()
    goal = epc_mw[matched].clip(lower=lo[matched], upper=hi[matched])
    return float(goal.sum() + lo[~matched].sum()), float(goal.sum() + hi[~matched].sum())

def balance_skips(SimAuto, final: dict[str, pd.DataFrame], targets: dict[str, pd.DataFrame] | None = None) -> set:
    """
    ObjectIDs a balance leaves alone: gens/loads on buses outside the main island (the swing can't take what they
    move) or outside the scaling voltage band (the balance's own scaling would hold them at once, so their share would
    never move), and, during scaling, those the main scaling holds back or has excluded (`targets`).
    """
    vpu, island = scale.bus_vpu(SimAuto), scale.bus_islands(SimAuto)
    if not (island == scale.MAIN_ISLAND).any():     # Numbering not as expected (or blank): don't skip everything.
        island = pd.Series(scale.MAIN_ISLAND, index=island.index)
    skip = {oid for df in final.values()
            for oid in df.loc[~scale.voltage_ok(df, vpu) | (df['BusNum'].map(island) != scale.MAIN_ISLAND), 'ObjectID']}
    for df in (targets or {}).values():
        skip |= set(df.index[~scale.included(df) | scale.held(df)])
    return skip

def balance_targets(final: dict[str, pd.DataFrame], swing: float, swing_buses, skip=frozenset()) -> dict[str, pd.DataFrame]:
    """
    Targets which move `swing` MW off the swing onto generation only; load is never moved (it stays on its EPC target).
    (Swing at +X MW: gens up X.) One factor on every Closed gen with MW > 0 (not on a swing bus, not a DC tie), so each
    area & GenType moves by the same share; what a group can't take (MWMax, or 0 going down) goes onto the rest of its
    area, then onto every gen (spread_evenly()). Statuses are unchanged. The Load table is empty.
    skip: ObjectIDs left out (balance_skips()); the rest take their share.
    """
    gen = scalable_gens(final['Gen'])
    gen = gen[closed(gen) & (gen['MWSetPoint'] > 0) & ~gen['BusNum'].isin(swing_buses) & ~gen['ObjectID'].isin(skip)].copy()
    cap = gen['MWMax'].where(gen['MWMax'] > gen['MWSetPoint'], gen['MWSetPoint'])  # A unit already above MWMax isn't raised.
    change, _ = spread_evenly(swing, gen['MWSetPoint'], cap, *gen_keys(gen))
    gen['MWSetPoint_Target'] = gen['MWSetPoint'] + change
    gen['MWSetPoint_Seed'] = gen['MWSetPoint']
    gen['Status_Target'] = gen['Status']

    load = final['Load'].iloc[0:0].copy()
    for c in scale.SCALED['Load'] + list(scale.STATUS['Load']):
        load[f'{c}_Target'] = load[c]
    load['P_Seed'] = pd.Series(dtype=float)

    targets = {'Gen': gen, 'Load': load}
    for df in targets.values():
        df['ExclusionReason'] = ''
    return {t: df.set_index('ObjectID') for t, df in targets.items()}

def prebalance_targets(targets: dict[str, pd.DataFrame], loss_change: float = 0.0, swing_change: float = 0.0) -> dict[str, pd.DataFrame] | None:
    """
    Before scaling: moves the targets so that scaling leaves the swing where it is. The change in
    load - gen - dist gen + losses from now to the targets would otherwise land on the swing as it scales (and set off a
    midrun balance every few steps). Gens and loads count only while Closed (an Open one injects nothing, whatever its MW).
    loss_change: estimated target losses less seed losses (loss_estimate()). swing_change: the case's own slack gens'
    seed MW less their EPC MW (swing_epc()): the rest of the gens make that up, so the slack ends at its EPC MW.
    Same rule as a balance: all onto gen, one
    factor on the target MW of every included gen Closed in the target with a positive target, spread_evenly() (capped
    at MWMax); load targets are never moved. What the gens can't take stays on the swing. Returns the moved tables for
    the `runs` sheet (mw_changes()), or None if nothing moved.
    """
    g, l = targets['Gen'], targets['Load']
    g_in, l_in = g[scale.included(g)], l[scale.included(l)]
    on = lambda df, suffix: df['Status' + suffix] == 'Closed'
    p = l_in[LOAD_P].sum(axis=1).where(on(l_in, ''), 0.0)
    p_target = l_in[[f'{c}_Target' for c in LOAD_P]].sum(axis=1).where(on(l_in, '_Target'), 0.0)
    gen_mw = lambda suffix: g_in['MWSetPoint' + suffix].where(on(g_in, suffix), 0.0).sum()
    def dist(suffix):       # Dist gen injects only while the load and its dist gen are both Closed.
        return l_in['DistMWInput' + suffix].where(on(l_in, suffix) & (l_in['DistStatus' + suffix] == 'Closed'), 0.0).sum()
    swing = ((p_target - p).sum() - (gen_mw('_Target') - gen_mw('')) - (dist('_Target') - dist(''))
             + loss_change + swing_change)
    if abs(swing) < 1.0:
        return None

    gen = g_in[(g_in['Status_Target'] == 'Closed') & (g_in['MWSetPoint_Target'] > 0)]
    seed = gen['MWSetPoint_Target']
    cap = gen['MWMax'].where(gen['MWMax'] > seed, seed)     # A target already above MWMax isn't raised.
    new = seed + spread_evenly(swing, seed, cap, *gen_keys(gen))[0] if len(gen) else seed
    g.loc[gen.index, 'MWSetPoint_Target'] = new

    left = swing - (new.sum() - seed.sum())
    print(f'Targets balanced before scaling: {swing:+.0f} MW would have landed on the swing'
          + (f' (losses {loss_change:+.0f} MW)' if abs(loss_change) >= 1.0 else '')
          + (f' (slack to its EPC MW {swing_change:+.0f} MW)' if abs(swing_change) >= 1.0 else '') + f'; gen {new.sum() - seed.sum():+.0f} MW'
          + (f'; {left:+.0f} MW could not be placed and stays on the swing.' if abs(left) >= 1.0 else '.'))
    return {'Gen': pd.DataFrame({'MWSetPoint': new, 'MWSetPoint_Seed': seed}),
            'Load': pd.DataFrame(columns=LOAD_P + ['P_Seed'], dtype=float)}

def mw_changes(moved: dict[str, pd.DataFrame]) -> dict[str, float]:
    """MW moved onto gens and loads by a balance (balance_targets() tables, after scaling)."""
    return {'Gen MW Change': (moved['Gen']['MWSetPoint'] - moved['Gen']['MWSetPoint_Seed']).sum(),
            'Load MW Change': (moved['Load'][LOAD_P].sum(axis=1) - moved['Load']['P_Seed']).sum()}

def balance_swing(SimAuto, scaled: dict[str, pd.DataFrame] | None = None, tables: list | None = None,
                  epc: pd.DataFrame | None = None) -> list[scale.Run]:
    """
    With the case open: moves the swing onto the whole case's gen (see balance_targets()) until it is within
    BALANCE_TOL_MW of its goal (swing_goal(): 0 for the giant swing; the case's own slack gens at their EPC MW where
    their key is in the EPC (`epc`, swing_epc()), else anywhere within their limits). Returns a scale.Run per round that ran
    (`balance<N>`; the swing after a round is the next round's, or the end swing).
    scaled: the main scaling's tables, so units it held back or excluded are left alone here too (balance_skips()).
    tables: if a list, each round's balance_targets() tables (after scaling) are appended to it, for areas().
    """
    slack = prep.slack_buses(SimAuto)
    runs: list[scale.Run] = []
    exhausted = True
    for i in range(BALANCE_MAX_ROUNDS):
        try:
            if not _balance_round(SimAuto, i, slack, epc, scaled, tables, runs):
                exhausted = False
                break
        except Exception as e:      # Keep the rounds that ran (their tables are in `tables`, with what each applied).
            print(f'WARNING: Final balance {i + 1} failed: {e!r}')
            if runs and runs[-1].name == f'balance{i + 1}':     # It failed inside its own scaling: note it on its row.
                runs[-1].note = '; '.join(filter(None, [runs[-1].note, f'failed: {e!r}']))
            else:
                runs.append(scale.Run(f'balance{i + 1}', note=f'failed: {e!r}'))
            exhausted = False
            break
    try:
        swing, _ = swing_mw(read_case(SimAuto)['Gen'], slack)
    except Exception:
        swing = np.nan
    for run, nxt in zip(runs, runs[1:]):
        run.after = nxt.before if pd.notna(nxt.before) else swing     # A failed round has no Before.
    if runs:
        runs[-1].after = swing
        if exhausted:
            runs[-1].note = '; '.join(filter(None, [runs[-1].note, 'stopped at max rounds']))
    return runs

def _balance_round(SimAuto, i, slack, epc, scaled, tables, runs) -> bool:
    """One balance_swing() round; False once the swing is at its goal or the round's scaling stopped."""
    final = read_case(SimAuto)
    swing, swing_buses = swing_mw(final['Gen'], slack)
    lo, hi = swing_goal(final['Gen'], slack, epc)
    print(f'----- Balancing round {i + 1}: swing at {swing:.1f} MW' + (f' (goal {lo:.0f} to {hi:.0f} MW)' if (lo, hi) != (0.0, 0.0) else '') + ' -----')
    if lo - BALANCE_TOL_MW < swing < hi + BALANCE_TOL_MW:
        return False
    m = min(BALANCE_TOL_MW, (hi - lo) / 2)
    aim = float(np.clip(swing, lo + m, hi - m))
    targets = balance_targets(final, swing - aim, swing_buses, balance_skips(SimAuto, final, scaled))
    if tables is not None:
        tables.append(targets)
    run = scale.Run(f'balance{i + 1}', moved=targets, before=swing)
    runs.append(run)
    try:
        run.log = scale.iterate_to_gen_load_targets(SimAuto, targets, label=f'Final balance {i + 1}', initial_step=scale.MAX_STEP)
    except pw.SimAutoError as e:
        # Keep the hour: the case is solved at the last good checkpoint, and `targets` holds what was applied.
        print(f'WARNING: Final balance {i + 1} stopped: {e}')
        run.log, run.note = getattr(e, 'log', {}), f'stopped: {e}'
        return False
    return True

def rescale_targets(targets: dict[str, pd.DataFrame], moved: dict[str, pd.DataFrame]):
    """
    After a balance (`moved`: balance_targets() tables, after scaling) in the middle of scaling to `targets`:
    works out the factor each moved row was moved by, and has scale.rescale() apply it to the rest of the scaling, so
    that keeps the balance instead of undoing it. Gen targets stay within MWMax.
    """
    gen = moved['Gen']
    scale.rescale(targets['Gen'], gen['MWSetPoint'] / gen['MWSetPoint_Seed'], scale.SCALED['Gen'], cap=gen['MWMax'])

# ------------------ Make-up ------------------
MAKEUP = True   # Keep area gen totals: plan Unmet and excluded units' shortfall go onto the group's, then the area's, other units.

def makeup_targets(final: dict[str, pd.DataFrame], targets: dict[str, pd.DataFrame], swing_buses,
                   skip=frozenset(), nonconforming=frozenset()) -> tuple[dict[str, pd.DataFrame], float] | None:
    """
    After the main scaling (area mode): per (AreaNumber, GenType) and per AreaNumber's load, the shortfall S = target -
    value of the main rows which were excluded (held units are excluded at the end of the main run; islanded rows are
    left out). S goes onto other units, read from the live case (`final`, read_case()), so the main tables are never
    touched. Gens: the same group's Closed gens with MW > 0 not on a swing bus, in proportion to MW (capped at MWMax
    going up, floor 0 going down); what they can't take goes onto the rest of the area's gens of every GenType, the same
    way (spread_by_group(), spill_all=False), and what the area can't take stays short. Loads: the area's Closed loads
    with MW > 0, one factor on MW and Mvar; the non-conforming loads (`nonconforming`: reports['nonconforming']) neither
    count as short nor take a share. skip: ObjectIDs left out (balance_skips()).
    Returns (tables in balance_targets()' shape, {'Gen': MW asked, 'Load': MW asked}), or None if no group is short
    by 1 MW or more.
    """
    def shortfall(df, cur, tgt, group):
        out = df[~scale.included(df) & ~scale.is_island_reason(df['ExclusionReason'])]
        return _by_group(out, tgt(out) - cur(out), group(out) if group else None)
    tp = lambda df: df[[f'{c}_Target' for c in LOAD_P]].sum(axis=1)
    nonc = set(nonconforming)
    s_gen = shortfall(targets['Gen'], lambda d: d['MWSetPoint'], lambda d: d['MWSetPoint_Target'], classify_gen_type)
    s_load = shortfall(targets['Load'][~targets['Load'].index.isin(nonc)], lambda d: d[LOAD_P].sum(axis=1), tp, None)
    s_gen, s_load = s_gen[s_gen.abs() >= 1.0], s_load[s_load.abs() >= 1.0]
    if not (len(s_gen) or len(s_load)):
        return None

    gen = scalable_gens(final['Gen'])
    gen = gen[closed(gen) & (gen['MWSetPoint'] > 0) & ~gen['BusNum'].isin(swing_buses) & ~gen['ObjectID'].isin(skip)
              & gen['AreaNumber'].isin(s_gen.index.get_level_values('AreaNumber'))].copy()
    cap = gen['MWMax'].where(gen['MWMax'] > gen['MWSetPoint'], gen['MWSetPoint'])
    change, _ = spread_by_group(s_gen, gen['MWSetPoint'], cap, *gen_keys(gen), spill_all=False)
    gen = gen[change.abs() > 1e-9]
    gen['MWSetPoint_Target'] = gen['MWSetPoint'] + change[gen.index]
    gen['MWSetPoint_Seed'] = gen['MWSetPoint']
    gen['Status_Target'] = gen['Status']

    load = final['Load']
    load = load[closed(load) & ~load['ObjectID'].isin(nonc | set(skip))].copy()
    p = load[LOAD_P].sum(axis=1)
    load = load[(p > 0) & load['AreaNumber'].isin(s_load.index.get_level_values('AreaNumber'))]
    p = p[load.index]
    k = (1 + load['AreaNumber'].map(s_load.droplevel('Group')) / load['AreaNumber'].map(p.groupby(load['AreaNumber']).sum())).clip(lower=0.0)
    for c in scale.SCALED['Load']:
        load[f'{c}_Target'] = load[c] * k if c in LOAD_P + LOAD_Q else load[c]
    for c in scale.STATUS['Load']:
        load[f'{c}_Target'] = load[c]
    load['P_Seed'] = p

    tables = {'Gen': gen, 'Load': load}
    for df in tables.values():
        df['ExclusionReason'] = ''
    return {t: df.set_index('ObjectID') for t, df in tables.items()}, {'Gen': float(s_gen.sum()), 'Load': float(s_load.sum())}

def makeup(SimAuto, targets: dict[str, pd.DataFrame], tables: list | None = None, nonconforming=frozenset()) -> list[scale.Run]:
    """
    With the case open, after the main scaling: runs makeup_targets() (a `makeup` run; none if nothing is short).
    tables: if a list, the make-up tables (after scaling) are appended to it, for areas(). A make-up which fails
    part-way keeps what it applied and the hour carries on.
    """
    slack = prep.slack_buses(SimAuto)
    final = read_case(SimAuto)
    before, swing_buses = swing_mw(final['Gen'], slack)
    made = makeup_targets(final, targets, swing_buses, balance_skips(SimAuto, final, targets), nonconforming)
    if made is None:
        return []
    moved, asked = made
    if tables is not None:
        tables.append(moved)
    run = scale.Run('makeup', moved=moved, before=before)
    try:
        run.log = scale.iterate_to_gen_load_targets(SimAuto, moved, label='Make-up', initial_step=scale.MAX_STEP)
    except pw.SimAutoError as e:
        print(f'WARNING: Make-up stopped: {e}')
        run.log, run.note = getattr(e, 'log', {}), f'stopped: {e}'
    run.after = swing_mw(read_case(SimAuto)['Gen'], slack)[0]
    placed = mw_changes(moved)
    run.note = '; '.join(filter(None, [f"gen asked {asked['Gen']:+,.0f} MW, placed {placed['Gen MW Change']:+,.0f} MW; "
                                       f"load asked {asked['Load']:+,.0f} MW, placed {placed['Load MW Change']:+,.0f} MW", run.note]))
    print(f'Make-up: {run.note}.')
    return [run]

def midrun_balancer(SimAuto, epc: pd.DataFrame | None = None):
    """
    Returns (after_step, runs) for scale.iterate_to_gen_load_targets(). After each step, if the swing is outside
    its band, moves it onto the whole case's gen (balance_targets()), and keeps that in the scaling still to come
    (rescale_targets()). The swing itself is never moved.
    - The giant swing: the band is swing_band(); a balance aims at 0 MW, or at the nearest point a quarter of the band
      inside its edges if 0 is not there.
    - The case's own slack: the pre-balance plans it to go from where it starts to its goal (swing_goal(), `epc`:
      swing_epc()) as the scaling goes, so at each step it is expected at that point of the way. The band is that
      point +/- BALANCE_TRIGGER_MW, within its MWMin / MWMax, but at least a half trigger either side of the point a
      balance aims at (so a zero-width range still has a band); a balance aims back at that point. With a trigger of 0
      the band is its MWMin / MWMax, and a range under MIN_OWN_RANGE_MW is not balanced midrun. A slack which strays far
      diverges (about 2 GW on the owner's case), so this keeps it near its plan.
    If a balance leaves the swing outside the band (e.g. gens at MWMax), the band is widened to where it got to, plus
    a margin (BALANCE_TRIGGER_MW, or 10% of the band if that's 0), so the next balance waits until the swing has moved
    that much further.
    Step size: after_step measures how fast the swing drifts from where it should be (per 1/RESOLUTION of progress)
    and returns the largest next step which keeps that drift within BALANCE_TRIGGER_MW, so one big step can't carry
    the swing far past its band before a balance sees it.
    runs: a scale.Run per balance (`midrun<N>`), filled in as they run.
    """
    slack = prep.slack_buses(SimAuto)
    gen0 = pw.get_df(SimAuto, 'Gen', GEN_FIELDS)
    lo, hi = swing_band(gen0, slack)
    units0 = swing_units(gen0, slack)
    own_slack = bool(len(units0)) and not (units0['ID'] == prep.SWING_ID).any()
    if own_slack:
        start = swing_mw(gen0, slack)[0]
        end = float(np.clip(start, *swing_goal(gen0, slack, epc)))     # Where the pre-balance planned it to end.
    trigger = BALANCE_TRIGGER_MW
    # With no trigger the own slack's band is its limits; a range under MIN_OWN_RANGE_MW would balance almost every step: off.
    own_off = own_slack and not trigger and hi - lo < MIN_OWN_RANGE_MW
    if own_off:
        print(f'Midrun balancing off for the case\'s own slack: balance_trigger_mw is 0 and its range ({lo:.0f} to {hi:.0f} MW) is under {MIN_OWN_RANGE_MW:.0f} MW.')
    if not own_slack and not (np.isfinite(lo) or np.isfinite(hi)):
        print('Midrun balancing off: balance_trigger_mw is 0 and balance_at_swing_limits is false.')

    def expected(progress):
        return float(np.clip(start + (end - start) * progress / scale.RESOLUTION, lo, hi))

    def band(progress):
        if not own_slack:
            return state['lo'], state['hi']
        if own_off:
            return -np.inf, np.inf
        ref = expected(progress)
        if not trigger:
            return lo - state['extra'], hi + state['extra']
        # Within the limits and +/- trigger of the plan, but at least trigger/2 either side of the aim (so a slack with a
        # zero or narrow range still gets a band to drift in).
        aim = aim_for(0.0, progress)
        return (min(max(lo, ref - trigger), aim - trigger / 2) - state['extra'],
                max(min(hi, ref + trigger), aim + trigger / 2) + state['extra'])

    # The case's own slack is aimed this far inside its limits at least, so a balance onto a planned point at a limit
    # (an EPC MW at its Pmax) leaves it room to drift before the next one, instead of a balance every 1% step.
    inside = min(BALANCE_TRIGGER_MW / 2 if BALANCE_TRIGGER_MW else np.inf, 0.25 * (hi - lo) if np.isfinite(hi - lo) else np.inf)
    inside = inside if np.isfinite(inside) else 0.0

    def aim_for(swing, progress):
        if own_slack:
            return float(np.clip(expected(progress), lo + inside, hi - inside))
        # Balance to a point inside the band, not to its edge, so the next small drift doesn't set off another balance:
        # 0, or with a swing whose MWMin is 0 (0 is the band's edge) a quarter of the band inside.
        if np.isfinite(lo) and np.isfinite(hi):
            return float(np.clip(0.0, lo + 0.25 * (hi - lo), hi - 0.25 * (hi - lo)))
        return float(np.clip(0.0, lo, hi))
    def ref(progress):     # Where the swing should be.
        return expected(progress) if own_slack else aim_for(0.0, progress)
    state = {'lo': lo, 'hi': hi, 'extra': 0.0, 'p': 0, 'dev': swing_mw(gen0, slack)[0] - ref(0)}
    runs: list[scale.Run] = []

    def step_cap(progress, swing_now, swing_after):
        # Drift per unit of progress since the last check, away from where the swing should be; the next step may
        # carry the swing at most to the edge of its band (as widened) in the direction it drifts.
        dev_now, dev_after = swing_now - ref(progress), swing_after - ref(progress)
        rate = (dev_now - state['dev']) / max(progress - state['p'], 1)
        state['p'], state['dev'] = progress, dev_after
        b_lo, b_hi = band(progress)
        room = (b_hi - swing_after) if rate > 0 else (swing_after - b_lo)
        if not np.isfinite(room) or abs(rate) < 1e-9:
            return None
        return max(1, int(room / abs(rate)))

    def after_step(SimAuto, targets, progress):
        gen = pw.get_df(SimAuto, 'Gen', GEN_FIELDS)
        swing, swing_buses = swing_mw(gen, slack)
        b_lo, b_hi = band(progress)
        if b_lo <= swing <= b_hi:
            return step_cap(progress, swing, swing)
        print(f'----- Swing at {swing:.1f} MW, outside {b_lo:.0f} to {b_hi:.0f} MW: balancing onto gen -----')
        final = {'Load': pw.get_df(SimAuto, 'Load', LOAD_FIELDS), 'Gen': gen}
        # Gens the scaling takes to 0 MW or below get no share: rescale_targets() would undo it.
        skip = balance_skips(SimAuto, final, targets) | set(targets['Gen'].index[targets['Gen']['MWSetPoint_Target'] <= 0])
        moved = balance_targets(final, swing - aim_for(swing, progress), swing_buses, skip)
        label = f'Balance {len(runs) + 1} (main scaling at {progress} of {scale.RESOLUTION})'
        run = scale.Run(f'midrun{len(runs) + 1}', moved=moved, before=swing, at=f'{progress} of {scale.RESOLUTION}')
        runs.append(run)
        notes = []
        try:
            run.log = scale.iterate_to_gen_load_targets(SimAuto, moved, label=label, initial_step=scale.MAX_STEP)
        except pw.SimAutoError as e:
            # Keep the hour: the case is solved at the last good checkpoint, and `moved` holds what was applied.
            print(f'WARNING: {label} stopped: {e}')
            run.log = getattr(e, 'log', {})
            notes.append(f'stopped: {e}')
        rescale_targets(targets, moved)
        after, _ = swing_mw(pw.get_df(SimAuto, 'Gen', GEN_FIELDS), slack)
        run.after = after
        if own_slack:
            state['extra'] = 0.0
            n_lo, n_hi = band(progress)
            if not n_lo <= after <= n_hi:
                margin = BALANCE_TRIGGER_MW or (0.1 * (hi - lo) if np.isfinite(hi - lo) else 0.0)
                state['extra'] = max(n_lo - after, after - n_hi) + margin
                notes.append(f'band widened by {state["extra"]:.0f} MW')
        elif not lo <= after <= hi:
            finite = np.isfinite(lo) and np.isfinite(hi)
            margin = (0.1 * (hi - lo) if finite else 0.0) if not BALANCE_TRIGGER_MW else BALANCE_TRIGGER_MW
            state['lo'], state['hi'] = min(lo, after - margin), max(hi, after + margin)
            notes.append(f'band widened to {state["lo"]:.0f}..{state["hi"]:.0f} MW')
        else:
            state['lo'], state['hi'] = lo, hi
        run.note = '; '.join(notes)
        return step_cap(progress, swing, after)

    return after_step, runs
