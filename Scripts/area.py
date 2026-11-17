"""
Step 02 targets: scale the PWB to the GridView hour by area, not element by element.

Topology is not matched: nothing is created or deleted, and no branch or load is opened or closed to match the
GridView hour. Gens are brought online or taken offline only to suit each bucket's target (unit commitment, below).
The targets are:
- Load, per AreaNumber: the EPC's total load MW and Mvar (constant power + current + impedance, at nominal voltage,
  Closed loads only; the EPC's distributed generation is ignored, as in pcm-to-pf).
  In the PWB, distributed generation is set Open & 0, constant current and impedance load go to 0, and the constant
  power MW / Mvar are scaled so the area total matches. Each load's starting share is its S + I + Z total.
  Positive and negative load are matched separately: each area's positive loads are scaled to the EPC's positive total
  for that area, and its negative loads (DC converters, generation modelled as load) to the EPC's negative total.
- Generation, per (AreaNumber, GenType): the EPC's total MWSetPoint of Closed gens of that type in that area.
  The PWB's Closed gens of that type in that area share it in proportion to MWMax (the same capacity factor on each),
  capped at MWMax. What can't be placed is reported as Unmet MW, and ends up on the giant swing.
  DC ties (unit type DC) are left alone in both cases: they are set by hand.
- Unit commitment (COMMIT): within each (area, GenType) bucket, gens are brought online or taken offline to suit the
  target (see commit()). Load statuses never change.
- Exception: an area whose list of loads (or gens) has exactly the same keys in both cases has each one targeted to its
  own EPC value (0 if Open in the EPC), instead of being scaled as a group. With COMMIT, such gens also take their EPC status.

After that, whatever is left on the swing is balanced out over the whole case, half on generation and half on load
(see balance_targets()), so the swing ends near 0 MW. This moves the area totals off their targets by that much.

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
SCALE_NONCONFORMING = False     # False: they keep their seed values, and are left out of both cases' area totals.

def conforming(load: pd.DataFrame) -> pd.Series:
    return load['AGC'].fillna('').astype(str).str.strip().str.upper() != 'NO'

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

def proportional_fill(target: float, weight: pd.Series, cap: pd.Series) -> pd.Series:
    """
    Spreads target (>= 0) over units in proportion to weight (>= 0), at most cap each: units which reach their cap stay
    there and the rest is spread again over the others. Units with weight 0 get 0. Less than target if all are capped.
    """
    x = pd.Series(0.0, index=weight.index)
    free = weight > 0
    remaining = target
    while free.any() and remaining > 1e-9:
        share = remaining * weight[free] / weight[free].sum()
        over = share > cap[free]
        if not over.any():
            x[free] = share
            break
        hit = over[over].index
        x[hit] = cap[hit]
        remaining -= cap[hit].sum()
        free[hit] = False
    return x

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
        mw = against + sign * proportional_fill(want, along, cap.where(cap > along, along))
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
    report['Unmet MW'] = report['Target MW'] - report['Planned MW']
    # Units offline before and after aren't touched.
    df = df[on_seed | on_plan]
    return df.set_index('ObjectID'), report.reset_index()

def build_targets(epc: dict[str, pd.DataFrame], case: dict[str, pd.DataFrame], swing_buses=()) -> tuple[dict, dict]:
    """
    Returns (targets {'Gen', 'Load'} for scale.iterate_to_gen_load_targets(), reports {'area_load', 'area_gen'}).
    Gens on swing_buses (the slack bus) are left out of both cases: the solution sets their MW, and swing balancing
    deals with what they carry.
    """
    epc_load, case_load = epc['Load'], case['Load']
    # Non-conforming is decided by key: a load which is non-conforming in the PWB, or flagged in the EPC, is out of both.
    nonc_ids = set(case_load.loc[~conforming(case_load), 'ObjectID'])
    epc_nonc = ~conforming(epc_load) | epc_load['ObjectID'].isin(nonc_ids)
    if not SCALE_NONCONFORMING:
        epc_load, case_load = epc_load[~epc_nonc], case_load[conforming(case_load)]
    load, load_report = load_targets(epc_load, case_load)
    nonc = lambda df, mask: df[mask & closed(df)].assign(P=lambda d: d[LOAD_P].sum(axis=1)).groupby('AreaNumber')['P'].sum()
    load_report['Seed Non-conforming MW'] = load_report['AreaNumber'].map(nonc(case['Load'], ~conforming(case['Load']))).fillna(0.0)
    load_report['EPC Non-conforming MW'] = load_report['AreaNumber'].map(nonc(epc['Load'], epc_nonc)).fillna(0.0)
    off_swing = lambda df: df[~df['BusNum'].isin(swing_buses)]
    gen, gen_report = gen_targets(off_swing(epc['Gen']), off_swing(case['Gen']))
    return {'Gen': gen, 'Load': load}, {'area_load': load_report, 'area_gen': gen_report}

def open_load_cleanup(case_load: pd.DataFrame) -> pd.DataFrame:
    """Open loads with distributed generation left on: set it Open & 0 (no effect on the solution)."""
    df = case_load[~closed(case_load) & (conforming(case_load) | SCALE_NONCONFORMING) & ((case_load['DistStatus'] == 'Closed') | (case_load['DistMWInput'] != 0) | (case_load['DistMvarInput'] != 0))]
    return df.assign(DistStatus='Open', DistMWInput=0.0, DistMvarInput=0.0)

AREAS_COLUMNS = ['AreaNumber', 'Group', 'Seed MW', 'Target MW', 'Unmet MW', 'Not Reached MW', 'Balance MW', 'Final MW', 'Error MW',
                 'Target Mvar', 'Final Mvar', 'Online Seed', 'Online Final', 'Note']
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

def _gen_group(df: pd.DataFrame) -> pd.Series:
    return classify_gen_type(df) if len(df) else pd.Series(dtype=object)

def areas(final: dict[str, pd.DataFrame] | None, reports: dict[str, pd.DataFrame], targets: dict[str, pd.DataFrame] | None = None,
          balance_tables=(), swing_buses=()) -> pd.DataFrame:
    """
    The `areas` sheet: a row per (AreaNumber, Group), Group a GenType or 'Load' (AREAS_COLUMNS). The ledger per row:
        Final = Target - Unmet - Not Reached + Balance
    Not Reached: over the main scaling's rows, end-of-run target less end-of-run value (excluded, diverged, held units).
    Balance: what swing balancing moved onto the group = the midrun part (end-of-run `_Target` less Planned MW, main rows:
    midrun balances rescale the main targets) + the final part (value less seed in each final balance's table).
    final: read_case() of the scaled case, or None before scaling (Final columns blank). targets: the main tables after
    scaling (None: nothing reached or balanced). balance_tables: balance_targets() tables of the final balance rounds,
    after scaling (balance_swing(..., tables=)). Gens on swing_buses are left out, as in build_targets().
    """
    key = ['AreaNumber', 'Group']
    plan = pd.concat([reports['area_gen'].rename(columns={'GenType': 'Group'}),
                      reports['area_load'].rename(columns={'Seed Loads': 'Seed Online'}).assign(Group='Load')]).set_index(key)
    parts: dict[str, pd.Series] = {}

    p_of = lambda df, cols: df[cols].sum(axis=1)
    if targets is not None:
        g, l = targets['Gen'], targets['Load']
        gg = _gen_group(g)
        not_reached = (_by_group(g, g['MWSetPoint_Target'] - g['MWSetPoint'], gg)
                       .add(_by_group(l, p_of(l, [f'{c}_Target' for c in LOAD_P]) - p_of(l, LOAD_P)), fill_value=0.0))
        end_target = (_by_group(g, g['MWSetPoint_Target'], gg).add(_by_group(l, p_of(l, [f'{c}_Target' for c in LOAD_P])), fill_value=0.0))
        planned = _by_group(g, g['Planned MW'], gg).add(_by_group(l, l['Planned MW']), fill_value=0.0)
        parts['Not Reached MW'] = not_reached
        parts['Balance MW'] = end_target.sub(planned, fill_value=0.0)        # The midrun part.
    for t in balance_tables:
        g, l = t['Gen'], t['Load']
        g = g[~g['BusNum'].isin(swing_buses)]       # Left out of Final MW too (a second slack bus beside the giant swing).
        parts['Balance MW'] = (parts.get('Balance MW', pd.Series(dtype=float))
                               .add(_by_group(g, g['MWSetPoint'] - g['MWSetPoint_Seed'], _gen_group(g)), fill_value=0.0)
                               .add(_by_group(l, p_of(l, LOAD_P) - l['P_Seed']), fill_value=0.0))
    if final is not None:
        load = final['Load'][closed(final['Load']) & (conforming(final['Load']) | SCALE_NONCONFORMING)]
        gen = scalable_gens(final['Gen'])
        gen = gen[closed(gen) & ~gen['BusNum'].isin(swing_buses)]
        parts['Final MW'] = _by_group(gen, gen['MWSetPoint'], _gen_group(gen)).add(_by_group(load, p_of(load, LOAD_P)), fill_value=0.0)
        parts['Final Mvar'] = _by_group(load, p_of(load, LOAD_Q))
        parts['Online Final'] = _by_group(gen, 1.0, _gen_group(gen)).add(_by_group(load, 1.0), fill_value=0.0)

    index = plan.index
    for s in parts.values():
        index = index.union(s.index)
    index = index.sort_values()
    out = pd.DataFrame(index=index)
    for c in ('Seed MW', 'Target MW', 'Unmet MW', 'Target Mvar', 'Seed Online'):
        out[c] = plan[c].reindex(index) if c in plan else np.nan
    out[['Seed MW', 'Target MW', 'Unmet MW']] = out[['Seed MW', 'Target MW', 'Unmet MW']].fillna(0.0)
    out.loc[out.index.get_level_values('Group') != 'Load', 'Target Mvar'] = np.nan
    for c in ('Not Reached MW', 'Balance MW'):
        out[c] = parts[c].reindex(index, fill_value=0.0) if c in parts else 0.0
    out = out.rename(columns={'Seed Online': 'Online Seed'})
    out['Online Seed'] = out['Online Seed'].fillna(0.0)
    for c in ('Final MW', 'Final Mvar', 'Online Final'):
        out[c] = parts[c].reindex(index, fill_value=0.0) if final is not None else np.nan
    out.loc[out.index.get_level_values('Group') != 'Load', 'Final Mvar'] = np.nan
    out['Error MW'] = out['Final MW'] - out['Target MW']
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Balance MW']
    out['Note'] = _area_notes(out, plan.reindex(index), ledger)
    return out.reset_index()[AREAS_COLUMNS]

def _area_notes(out: pd.DataFrame, plan: pd.DataFrame, ledger: pd.Series) -> list[str]:
    notes = []
    for i, k in enumerate(out.index):
        o, p, n = out.iloc[i], plan.iloc[i], []
        target, unmet = o['Target MW'], o['Unmet MW']
        if k[1] == 'Load':
            if target != 0 and p.get('Seed Online', 0) == 0:
                n.append('no PWB units')
            nonc = p.get('Seed Non-conforming MW', 0)
            if pd.notna(nonc) and nonc != 0:
                n.append(f'non-conforming left as is: {nonc:.1f} MW')
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
    SimAuto.CloseCase()
    return build_targets(epc, case, swing_buses=slack)

# ------------------ Swing balancing ------------------
BALANCE_TOL_MW = 1.0        # Stop balancing once the swing is within this of 0 MW...
BALANCE_MAX_ROUNDS = 10     # ...or after this many rounds (losses move with each round, so it takes a few).
# While scaling (midrun_balancer()), balance once the swing goes outside its band:
BALANCE_TRIGGER_MW = 1000.0         # further than this from 0 MW (0: not used)...
BALANCE_AT_SWING_LIMITS = False     # ...or past the swing unit's own MWMax / MWMin (summed, if several).


def swing_units(gen: pd.DataFrame, slack_buses) -> pd.DataFrame:
    """The giant swing (ID prep.SWING_ID) if the case has one Closed, else every Closed gen on a slack bus."""
    on = gen[closed(gen)]
    swing = on[on['ID'] == prep.SWING_ID]
    return swing if len(swing) else on[on['BusNum'].isin(slack_buses)]

def swing_mw(gen: pd.DataFrame, slack_buses) -> tuple[float, list]:
    """MW on the swing, and the swing's bus numbers."""
    swing = swing_units(gen, slack_buses)
    return float(swing['MW'].sum()), sorted(set(swing['BusNum']))

def swing_band(gen: pd.DataFrame, slack_buses) -> tuple[float, float]:
    """
    (low, high) MW the swing may sit within while scaling, before midrun_balancer() rebalances: the narrowest of
    +/- BALANCE_TRIGGER_MW (if not 0) and the swing's MWMin / MWMax (if BALANCE_AT_SWING_LIMITS).
    """
    lo, hi = -np.inf, np.inf
    if BALANCE_TRIGGER_MW:
        lo, hi = -BALANCE_TRIGGER_MW, BALANCE_TRIGGER_MW
    if BALANCE_AT_SWING_LIMITS:
        swing = swing_units(gen, slack_buses)
        lo, hi = max(lo, float(swing['MWMin'].sum())), min(hi, float(swing['MWMax'].sum()))
    return lo, hi

def balance_skips(SimAuto, final: dict[str, pd.DataFrame], targets: dict[str, pd.DataFrame] | None = None) -> set:
    """
    ObjectIDs a balance leaves alone: gens/loads on buses outside the scaling voltage band (the balance's own scaling
    would hold them at once, so their share would never move), and, during scaling, those the main scaling holds back
    or has excluded (`targets`).
    """
    vpu = scale.bus_vpu(SimAuto)
    skip = {oid for df in final.values() for oid in df.loc[~scale.voltage_ok(df, vpu), 'ObjectID']}
    for df in (targets or {}).values():
        skip |= set(df.index[~scale.included(df) | scale.held(df)])
    return skip

def balance_targets(final: dict[str, pd.DataFrame], swing: float, swing_buses, skip=frozenset()) -> dict[str, pd.DataFrame]:
    """
    Targets which move `swing` MW off the swing: half onto generation, half onto load, over the whole case.
    (Swing at +X MW: gens up X/2, load down X/2.) One factor on every Closed gen with MW > 0 (not on a swing bus, not a
    DC tie), capped at MWMax, and one on the MW and Mvar of every Closed load with MW > 0. Statuses are unchanged.
    skip: ObjectIDs left out (balance_skips()); the rest take their share.
    """
    gen = scalable_gens(final['Gen'])
    gen = gen[closed(gen) & (gen['MWSetPoint'] > 0) & ~gen['BusNum'].isin(swing_buses) & ~gen['ObjectID'].isin(skip)].copy()
    cap = gen['MWMax'].where(gen['MWMax'] > gen['MWSetPoint'], gen['MWSetPoint'])  # A unit already above MWMax isn't raised.
    # In proportion to MW; what units at MWMax can't take is spread over the rest, so one round moves it all.
    gen['MWSetPoint_Target'] = proportional_fill(max(gen['MWSetPoint'].sum() + swing / 2, 0.0), gen['MWSetPoint'], cap)
    gen['MWSetPoint_Seed'] = gen['MWSetPoint']
    gen['Status_Target'] = gen['Status']

    load = final['Load']
    load = load[closed(load) & (conforming(load) | SCALE_NONCONFORMING) & ~load['ObjectID'].isin(skip)].copy()
    p = load[LOAD_P].sum(axis=1)
    load = load[p > 0]
    k_load = 1 - (swing / 2) / p[p > 0].sum() if len(load) else 1.0
    for c in scale.SCALED['Load']:
        load[f'{c}_Target'] = load[c] * max(k_load, 0.0) if c in LOAD_P + LOAD_Q else load[c]
    for c in scale.STATUS['Load']:
        load[f'{c}_Target'] = load[c]
    load['P_Seed'] = p[p > 0]

    targets = {'Gen': gen, 'Load': load}
    for df in targets.values():
        df['ExclusionReason'] = ''
    return {t: df.set_index('ObjectID') for t, df in targets.items()}

def prebalance_targets(targets: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame] | None:
    """
    Before scaling: moves the targets so that, losses aside, scaling leaves the swing where it is. The change in
    gen + dist gen - load from now to the targets would otherwise land on the swing as it scales (and set off a midrun
    balance every few steps). Same rule as a balance: half onto gen (in proportion to target MW, capped at MWMax, the
    rest onto load), half onto load (one factor on its MW and Mvar targets). Only included rows Closed in the target
    with a positive target move. Returns the moved tables for the `runs` sheet (mw_changes()), or None if nothing moved.
    """
    g, l = targets['Gen'], targets['Load']
    g_in, l_in = g[scale.included(g)], l[scale.included(l)]
    p, p_target = l_in[LOAD_P].sum(axis=1), l_in[[f'{c}_Target' for c in LOAD_P]].sum(axis=1)
    def dist(suffix):       # Dist gen injects only while the load and its dist gen are both Closed.
        on = (l_in['Status' + suffix] == 'Closed') & (l_in['DistStatus' + suffix] == 'Closed')
        return l_in['DistMWInput' + suffix].where(on, 0.0).sum()
    swing = (p_target - p).sum() - (g_in['MWSetPoint_Target'] - g_in['MWSetPoint']).sum() - (dist('_Target') - dist(''))
    if abs(swing) < 1.0:
        return None

    gen = g_in[(g_in['Status_Target'] == 'Closed') & (g_in['MWSetPoint_Target'] > 0)]
    seed = gen['MWSetPoint_Target']
    cap = gen['MWMax'] if 'MWMax' in gen else seed         # No MWMax known (topology): gens aren't raised.
    cap = cap.where(cap > seed, seed)                       # A target already above MWMax isn't raised.
    new = proportional_fill(max(seed.sum() + swing / 2, 0.0), seed, cap) if len(gen) else seed
    g.loc[gen.index, 'MWSetPoint_Target'] = new

    load = l_in[(l_in['Status_Target'] == 'Closed') & (p_target > 0)]
    load_p = p_target[load.index]
    load_down = swing - (new.sum() - seed.sum())            # What the gens couldn't take goes onto load too.
    k = max(1 - load_down / load_p.sum(), 0.0) if len(load) else 1.0
    for c in LOAD_P + LOAD_Q:
        l.loc[load.index, f'{c}_Target'] = l.loc[load.index, f'{c}_Target'] * k

    left = swing - (new.sum() - seed.sum()) + load_p.sum() * (k - 1)
    print(f'Targets balanced before scaling: {swing:+.0f} MW would have landed on the swing; '
          f'gen {new.sum() - seed.sum():+.0f} MW, load {load_p.sum() * (k - 1):+.0f} MW'
          + (f'; {left:+.0f} MW could not be placed and stays on the swing.' if abs(left) >= 1.0 else '.'))
    return {'Gen': pd.DataFrame({'MWSetPoint': new, 'MWSetPoint_Seed': seed}),
            'Load': load_p.mul(k).to_frame('SMW').assign(IMW=0.0, ZMW=0.0, P_Seed=load_p)}

def mw_changes(moved: dict[str, pd.DataFrame]) -> dict[str, float]:
    """MW moved onto gens and loads by a balance (balance_targets() tables, after scaling)."""
    return {'Gen MW Change': (moved['Gen']['MWSetPoint'] - moved['Gen']['MWSetPoint_Seed']).sum(),
            'Load MW Change': (moved['Load'][LOAD_P].sum(axis=1) - moved['Load']['P_Seed']).sum()}

def balance_swing(SimAuto, scaled: dict[str, pd.DataFrame] | None = None, tables: list | None = None) -> list[scale.Run]:
    """
    With the case open: moves what is left on the swing onto the whole case's gen and load, half each (see
    balance_targets()), round after round until it is within BALANCE_TOL_MW. Returns a scale.Run per round that ran
    (`balance<N>`; the swing after a round is the next round's, or the end swing).
    scaled: the main scaling's tables, so units it held back or excluded are left alone here too (balance_skips()).
    tables: if a list, each round's balance_targets() tables (after scaling) are appended to it, for areas().
    """
    slack = prep.slack_buses(SimAuto)
    runs: list[scale.Run] = []
    exhausted = True
    for i in range(BALANCE_MAX_ROUNDS):
        final = read_case(SimAuto)
        swing, swing_buses = swing_mw(final['Gen'], slack)
        print(f'----- Balancing round {i + 1}: swing at {swing:.1f} MW -----')
        if abs(swing) < BALANCE_TOL_MW:
            exhausted = False
            break
        targets = balance_targets(final, swing, swing_buses, balance_skips(SimAuto, final, scaled))
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
            exhausted = False
            break
    swing, _ = swing_mw(read_case(SimAuto)['Gen'], slack)
    for run, nxt in zip(runs, runs[1:]):
        run.after = nxt.before
    if runs:
        runs[-1].after = swing
        if exhausted:
            runs[-1].note = '; '.join(filter(None, [runs[-1].note, 'stopped at max rounds']))
    return runs

def rescale_targets(targets: dict[str, pd.DataFrame], moved: dict[str, pd.DataFrame]):
    """
    After a balance (`moved`: balance_targets() tables, after scaling) in the middle of scaling to `targets`:
    works out the factor each moved row was moved by, and has scale.rescale() apply it to the rest of the scaling, so
    that keeps the balance instead of undoing it. Gen targets stay within MWMax.
    """
    gen, load = moved['Gen'], moved['Load']
    scale.rescale(targets['Gen'], gen['MWSetPoint'] / gen['MWSetPoint_Seed'], scale.SCALED['Gen'], cap=gen['MWMax'])
    scale.rescale(targets['Load'], load[LOAD_P].sum(axis=1) / load['P_Seed'], LOAD_P + LOAD_Q)

def midrun_balancer(SimAuto):
    """
    Returns (after_step, runs) for scale.iterate_to_gen_load_targets(). After each step, if the swing is outside
    its band (swing_band()), moves it onto the whole case's gen and load, half each (balance_targets()), and keeps that in
    the scaling still to come (rescale_targets()). If a balance leaves the swing outside the band (e.g. gens at MWMax),
    the band is widened to where it got to, plus a margin (BALANCE_TRIGGER_MW, or 10% of the band if that's 0), so the
    next balance waits until the swing has moved that much further. The swing itself is never moved.
    A balance aims the swing at 0 MW, or at the nearest point a quarter of the band inside its edges if 0 is not there.
    runs: a scale.Run per balance (`midrun<N>`), filled in as they run.
    """
    slack = prep.slack_buses(SimAuto)
    lo, hi = swing_band(pw.get_df(SimAuto, 'Gen', GEN_FIELDS), slack)
    if not (np.isfinite(lo) or np.isfinite(hi)):
        print('Midrun balancing off: balance_trigger_mw is 0 and balance_at_swing_limits is false.')
    # Balance to a point inside the band, not to 0: with a swing whose MWMin is 0, 0 is the band's edge, and the next
    # step down would set off another balance. The final balance still takes the swing to 0.
    if np.isfinite(lo) and np.isfinite(hi):
        aim = float(np.clip(0.0, lo + 0.25 * (hi - lo), hi - 0.25 * (hi - lo)))
    else:
        aim = float(np.clip(0.0, lo, hi))
    state = {'lo': lo, 'hi': hi}
    runs: list[scale.Run] = []

    def after_step(SimAuto, targets, progress):
        gen = pw.get_df(SimAuto, 'Gen', GEN_FIELDS)
        swing, swing_buses = swing_mw(gen, slack)
        if state['lo'] <= swing <= state['hi']:
            return
        print(f'----- Swing at {swing:.1f} MW, outside {state["lo"]:.0f} to {state["hi"]:.0f} MW: balancing half on gen, half on load -----')
        final = {'Load': pw.get_df(SimAuto, 'Load', LOAD_FIELDS), 'Gen': gen}
        moved = balance_targets(final, swing - aim, swing_buses, balance_skips(SimAuto, final, targets))
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
        if not lo <= after <= hi:
            finite = np.isfinite(lo) and np.isfinite(hi)
            margin = BALANCE_TRIGGER_MW or (0.1 * (hi - lo) if finite else 0.0)
            state['lo'], state['hi'] = min(lo, after - margin), max(hi, after + margin)
            notes.append(f'band widened to {state["lo"]:.0f}..{state["hi"]:.0f} MW')
        else:
            state['lo'], state['hi'] = lo, hi
        run.note = '; '.join(notes)

    return after_step, runs
