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
LOAD_FIELDS: dict[str, type] = {**_KEY, 'Status': str, 'DistStatus': str, **{c: float for c in scale.SCALED['Load']}}
GEN_FIELDS: dict[str, type] = {**_KEY, 'Status': str, 'MW': float, 'MWSetPoint': float, 'MWMax': float, 'MWMin': float,
                               'UnitTypeCode': str, 'FuelTypeCode': str}

LOAD_P = ['SMW', 'IMW', 'ZMW']
LOAD_Q = ['SMvar', 'IMvar', 'ZMvar']

def read_case(SimAuto) -> dict[str, pd.DataFrame]:
    """Loads & gens of the open case (PWB, or an EPC opened in PowerWorld)."""
    return {'Load': pw.get_df(SimAuto, 'Load', LOAD_FIELDS), 'Gen': pw.get_df(SimAuto, 'Gen', GEN_FIELDS)}

def same_keys(epc_df: pd.DataFrame, case_df: pd.DataFrame) -> set:
    """Areas whose elements have exactly the same ObjectIDs in both cases."""
    e = epc_df.groupby('AreaNumber')['ObjectID'].agg(frozenset)
    c = case_df.groupby('AreaNumber')['ObjectID'].agg(frozenset)
    return {a for a in c.index if a in e.index and e[a] == c[a]}

def closed(df: pd.DataFrame) -> pd.Series:
    return df['Status'] == 'Closed'

# ------------------ Load targets ------------------
def sign_totals(value: pd.Series, area: pd.Series, name: str) -> pd.DataFrame:
    """Per area: total of the positive values ('<name> +') and of the negative values ('<name> -')."""
    return pd.DataFrame({f'{name} +': value.clip(lower=0).groupby(area).sum(),
                         f'{name} -': value.clip(upper=0).groupby(area).sum()})

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

def load_targets(epc_load: pd.DataFrame, case_load: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Returns (targets, area report). targets holds the case's Closed loads. Open loads aren't scaled.
    """
    epc = epc_load.assign(P=epc_load[LOAD_P].sum(axis=1).where(closed(epc_load), 0.0),
                          Q=epc_load[LOAD_Q].sum(axis=1).where(closed(epc_load), 0.0))
    matched = same_keys(epc_load, case_load)

    df = case_load[closed(case_load)].copy()
    df['P'] = df[LOAD_P].sum(axis=1)
    df['Q'] = df[LOAD_Q].sum(axis=1)
    df['Method'] = np.where(df['AreaNumber'].isin(matched), 'Matched keys', 'Area scale')

    # Area scale: positive and negative load separately, for MW and for Mvar. See sign_share().
    tgt = pd.concat([sign_totals(epc['P'], epc['AreaNumber'], 'MW'), sign_totals(epc['Q'], epc['AreaNumber'], 'Mvar')], axis=1)
    seed = pd.concat([sign_totals(df['P'], df['AreaNumber'], 'MW'), sign_totals(df['Q'], df['AreaNumber'], 'Mvar')], axis=1)
    df['SMW_Target'] = sign_share(df['P'], df['AreaNumber'], tgt['MW +'], tgt['MW -'])
    # Mvar of a sign which no seed load in the area has is spread over its zero-Mvar loads in proportion to |MW|.
    df['SMvar_Target'] = sign_share(df['Q'], df['AreaNumber'], tgt['Mvar +'], tgt['Mvar -'], fallback=df['P'].abs())

    # Matched keys: each load to its own EPC value.
    m = df['Method'] == 'Matched keys'
    epc_by_id = epc.set_index('ObjectID')
    df['SMW_Target'] = df['SMW_Target'].where(~m, df['ObjectID'].map(epc_by_id['P']))
    df['SMvar_Target'] = df['SMvar_Target'].where(~m, df['ObjectID'].map(epc_by_id['Q']))

    for c in ['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']:
        df[f'{c}_Target'] = 0.0
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
        'Planned MW': df.groupby('AreaNumber')['SMW_Target'].sum(),
        'Planned Mvar': df.groupby('AreaNumber')['SMvar_Target'].sum(),
        **{f'Target {c}': tgt[c] for c in tgt}, **{f'Seed {c}': seed[c] for c in seed},
        **{f'Planned {c}': planned[c] for c in planned},
        'Seed Loads': df.groupby('AreaNumber').size(),
        'EPC Loads': epc[closed(epc)].groupby('AreaNumber').size(),
        # e.g. GridView's X1/X2 distributed generation loads, which aren't in the PWB.
        'EPC-only MW': epc[~epc['ObjectID'].isin(case_load['ObjectID'])].groupby('AreaNumber')['P'].sum(),
    }).fillna(0.0)
    report['Method'] = np.where(report.index.isin(matched), 'Matched keys', 'Area scale')
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

def commit(rows: pd.DataFrame, target: float, gen_type: str, epc_online: set) -> pd.Series:
    """
    Which units of one bucket are online (True) for `target` MW. Only units with capacity in the direction of the
    target count (MWMax > 0, or MWMin < 0 for a negative target); units without are left as they are (NaN).
    """
    cap = (rows['MWMax'] if target >= 0 else -rows['MWMin']).fillna(0.0).clip(lower=0)
    able = cap > 0
    online = pd.Series(np.nan, index=rows.index, dtype=object)
    online[able] = closed(rows)[able]
    if gen_type in ALL_ON_TYPES:
        if target != 0:
            online[able] = True
        elif gen_type in OFF_AT_ZERO_TYPES:
            online[able] = False
        return online

    need = abs(target) / COMMIT_MAX_CF
    on = online[able].astype(bool)
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
                break
            online[i] = False
            have -= cap[i]
    return online

def gen_targets(epc_gen: pd.DataFrame, case_gen: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Returns (targets, bucket report). targets holds the case's gens which are online now or after commitment,
    except the giant swing and DC ties.
    """
    epc_gen, case_gen = scalable_gens(epc_gen), scalable_gens(case_gen)
    epc = epc_gen.assign(GenType=classify_gen_type(epc_gen), P=epc_gen['MWSetPoint'].fillna(0.0).where(closed(epc_gen), 0.0))
    epc_online = set(epc.loc[closed(epc), 'ObjectID'])
    matched = same_keys(epc_gen, case_gen)

    df = case_gen.copy()
    df['GenType'] = classify_gen_type(df)
    df['Method'] = np.where(df['AreaNumber'].isin(matched), 'Matched keys', 'Area & type scale')
    df['Status_Target'] = df['Status']
    df['MWSetPoint_Target'] = df['MWSetPoint'].where(closed(df), 0.0)

    keys = ['AreaNumber', 'GenType']
    tgt = epc.groupby(keys)['P'].sum()
    for bucket, rows in df[df['Method'] != 'Matched keys'].groupby(keys):
        t = float(tgt.get(bucket, 0.0))
        if COMMIT:
            online = commit(rows, t, bucket[1], epc_online)
            decided = online.notna()
            df.loc[online.index[decided], 'Status_Target'] = np.where(online[decided].astype(bool), 'Closed', 'Open')
        on = df.loc[rows.index, 'Status_Target'] == 'Closed'
        x = pd.Series(0.0, index=rows.index)
        if t >= 0:   # Share by MWMax.
            x[on] = capacity_share(t, rows.loc[on, 'MWMax'].fillna(0.0).clip(lower=0))
        else:        # Net charging (storage): share by |MWMin|.
            x[on] = -capacity_share(-t, (-rows.loc[on, 'MWMin'].fillna(0.0)).clip(lower=0))
        df.loc[rows.index, 'MWSetPoint_Target'] = x

    # Matched keys: each unit to its own EPC status (when committing) and value.
    m = df['Method'] == 'Matched keys'
    by_id = epc.set_index('ObjectID')
    if COMMIT:
        df['Status_Target'] = df['Status_Target'].where(~m, df['ObjectID'].map(by_id['Status']))
    df['MWSetPoint_Target'] = df['MWSetPoint_Target'].where(~m, df['ObjectID'].map(by_id['P'])).where(df['Status_Target'] == 'Closed', 0.0)

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
        'EPC Online': epc[closed(epc)].groupby(keys).size(),
        'EPC Online MWMax': epc[closed(epc)].groupby(keys)['MWMax'].sum(),
    }).fillna(0.0)
    report = report[(report.drop(columns='All Units MWMax') != 0).any(axis=1)]   # Drop buckets with nothing online anywhere.
    report['Method'] = np.where(report.index.get_level_values('AreaNumber').isin(matched), 'Matched keys', 'Area & type scale')
    report['Unmet MW'] = report['Target MW'] - report['Planned MW']
    # Units offline before and after aren't touched.
    df = df[on_seed | on_plan]
    return df.set_index('ObjectID'), report.reset_index()

def build_targets(epc: dict[str, pd.DataFrame], case: dict[str, pd.DataFrame], swing_buses=()) -> tuple[dict, dict]:
    """
    Returns (targets {'Gen', 'Load'} for scale.iterate_to_gen_load_targets(), reports {'area_load', 'area_gen'}).
    The case's gens on swing_buses (the slack bus) aren't scaled: the solution sets their MW.
    """
    load, load_report = load_targets(epc['Load'], case['Load'])
    gen, gen_report = gen_targets(epc['Gen'], case['Gen'][~case['Gen']['BusNum'].isin(swing_buses)])
    return {'Gen': gen, 'Load': load}, {'area_load': load_report, 'area_gen': gen_report}

def open_load_cleanup(case_load: pd.DataFrame) -> pd.DataFrame:
    """Open loads with distributed generation left on: set it Open & 0 (no effect on the solution)."""
    df = case_load[~closed(case_load) & ((case_load['DistStatus'] == 'Closed') | (case_load['DistMWInput'] != 0) | (case_load['DistMvarInput'] != 0))]
    return df.assign(DistStatus='Open', DistMWInput=0.0, DistMvarInput=0.0)

def achieved(final: dict[str, pd.DataFrame], reports: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Adds what the scaled case holds (`final`: read_case() of it, after scaling) to the area reports."""
    load = final['Load'][closed(final['Load'])]
    got_p = load[LOAD_P].sum(axis=1).groupby(load['AreaNumber']).sum()
    got_q = load[LOAD_Q].sum(axis=1).groupby(load['AreaNumber']).sum()
    area_load = reports['area_load'].assign(**{
        'Final MW': lambda r: r['AreaNumber'].map(got_p).fillna(0.0),
        'Final Mvar': lambda r: r['AreaNumber'].map(got_q).fillna(0.0)})
    area_load['Final Error MW'] = area_load['Final MW'] - area_load['Target MW']
    area_load['Final Error Mvar'] = area_load['Final Mvar'] - area_load['Target Mvar']
    final_split = pd.concat([sign_totals(load[LOAD_P].sum(axis=1), load['AreaNumber'], 'MW'),
                             sign_totals(load[LOAD_Q].sum(axis=1), load['AreaNumber'], 'Mvar')], axis=1)
    for c in final_split:
        area_load[f'Final {c}'] = area_load['AreaNumber'].map(final_split[c]).fillna(0.0)
        area_load[f'Final Error {c}'] = area_load[f'Final {c}'] - area_load[f'Target {c}']

    gen = scalable_gens(final['Gen'])
    gen = gen[closed(gen)]
    got = gen.groupby(['AreaNumber', classify_gen_type(gen).rename('GenType')])['MWSetPoint'].sum().rename('Final MW')
    area_gen = reports['area_gen'].merge(got.reset_index(), on=['AreaNumber', 'GenType'], how='outer')
    area_gen = area_gen.fillna({c: 0.0 for c in area_gen.columns if c not in ('AreaNumber', 'GenType', 'Method')})
    area_gen['Final Error MW'] = area_gen['Final MW'] - area_gen['Target MW']
    return {'area_load': area_load, 'area_gen': area_gen}

# ------------------ Swing balancing ------------------
BALANCE_TOL_MW = 1.0        # Stop balancing once the swing is within this of 0 MW...
BALANCE_MAX_ROUNDS = 10     # ...or after this many rounds (losses move with each round, so it takes a few).

def swing_mw(gen: pd.DataFrame, slack_buses) -> tuple[float, list]:
    """
    MW on the swing, and the swing's bus numbers. The swing is the giant swing (ID prep.SWING_ID) if the case has one
    Closed, else every Closed gen on a slack bus.
    """
    on = gen[closed(gen)]
    swing = on[on['ID'] == prep.SWING_ID]
    if len(swing) == 0:
        swing = on[on['BusNum'].isin(slack_buses)]
    return float(swing['MW'].sum()), sorted(set(swing['BusNum']))

def balance_targets(final: dict[str, pd.DataFrame], swing: float, swing_buses) -> dict[str, pd.DataFrame]:
    """
    Targets which move `swing` MW off the swing: half onto generation, half onto load, over the whole case.
    (Swing at +X MW: gens up X/2, load down X/2.) One factor on every Closed gen with MW > 0 (not on a swing bus, not a
    DC tie), capped at MWMax, and one on the MW and Mvar of every Closed load with MW > 0. Statuses are unchanged.
    """
    gen = scalable_gens(final['Gen'])
    gen = gen[closed(gen) & (gen['MWSetPoint'] > 0) & ~gen['BusNum'].isin(swing_buses)].copy()
    k_gen = 1 + (swing / 2) / gen['MWSetPoint'].sum() if len(gen) else 1.0
    cap = gen['MWMax'].where(gen['MWMax'] > gen['MWSetPoint'], gen['MWSetPoint'])  # A unit already above MWMax isn't raised.
    gen['MWSetPoint_Target'] = (gen['MWSetPoint'] * max(k_gen, 0.0)).clip(upper=cap)
    gen['MWSetPoint_Seed'] = gen['MWSetPoint']
    gen['Status_Target'] = gen['Status']

    load = final['Load']
    load = load[closed(load)].copy()
    p = load[LOAD_P].sum(axis=1)
    load = load[p > 0]
    k_load = 1 - (swing / 2) / p[p > 0].sum() if len(load) else 1.0
    for c in scale.SCALED['Load']:
        load[f'{c}_Target'] = load[c] * max(k_load, 0.0) if c in LOAD_P + LOAD_Q else load[c]
    for c in scale.STATUS['Load']:
        load[f'{c}_Target'] = load[c]
    load['SMW_Seed'] = load['SMW']

    targets = {'Gen': gen, 'Load': load}
    for df in targets.values():
        df['ExclusionReason'] = ''
    return {t: df.set_index('ObjectID') for t, df in targets.items()}

def slack_buses(SimAuto) -> list[int]:
    return pw.get_df(SimAuto, 'Bus', {'Number': int, 'Slack': str}).query("Slack == 'YES'")['Number'].tolist()

def balance_swing(SimAuto) -> dict[str, pd.DataFrame]:
    """
    With the case open: moves what is left on the swing onto the whole case's gen and load, half each (see
    balance_targets()), round after round until it is within BALANCE_TOL_MW. Returns the log sheets.
    """
    slack = slack_buses(SimAuto)
    rounds, sheets = [], {}
    for i in range(BALANCE_MAX_ROUNDS):
        final = read_case(SimAuto)
        swing, swing_buses = swing_mw(final['Gen'], slack)
        print(f'----- Balancing round {i + 1}: swing at {swing:.1f} MW -----')
        if abs(swing) < BALANCE_TOL_MW:
            break
        targets = balance_targets(final, swing, swing_buses)
        log = scale.iterate_to_gen_load_targets(SimAuto, targets)
        rounds.append({'Round': i + 1, 'Swing MW': swing, 'Swing Buses': str(swing_buses),
                       'Gen MW Change': (targets['Gen']['MWSetPoint'] - targets['Gen']['MWSetPoint_Seed']).sum(),
                       'Load MW Change': (targets['Load']['SMW'] - targets['Load']['SMW_Seed']).sum()})
        sheets.update({f'balance{i + 1}_{name}': df for name, df in log.items() if len(df)})
    else:
        swing, _ = swing_mw(read_case(SimAuto)['Gen'], slack)
    rounds.append({'Round': 'End', 'Swing MW': swing})
    return {'balance_df': pd.DataFrame(rounds), **sheets}
