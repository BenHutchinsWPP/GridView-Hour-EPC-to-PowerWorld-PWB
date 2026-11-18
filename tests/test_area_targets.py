"""
Tests for area.py targets, plan, commit, dispatch, key basis and non-conforming loads. Needs no PowerWorld.
    python tests/test_area_targets.py      (all tests: python tests/run.py [name part])
"""
import contextlib
import io
import numpy as np
import pandas as pd
from helpers import FakeSimAuto, SwingSimAuto, gens, loads, results, saved_settings
from Scripts import scale, area, topo

def test_classify_gen_type():
    df = pd.DataFrame({
        'UnitTypeCode': ['HY', 'UN', 'PV (Photovoltaic)', 'OT', '', 'SC', 'ZZ', 'BA', 'w2', None],
        'FuelTypeCode': ['WAT', 'NG', '', 'SUN', '', 'NG', 'WND', 'MWH', '', None]})
    assert area.classify_gen_type(df).tolist() == [
        'Hydro', 'Thermal', 'Solar', 'Solar', 'Unknown', 'Unknown', 'Wind', 'Battery Storage', 'Wind', 'Unknown']

def test_load_targets():
    case = loads([('l1', 1, 'Closed', 60.0, 20.0, 20.0, 10.0), ('l2', 1, 'Closed', 100.0, 0.0, 0.0, 10.0),
                  ('lo', 1, 'Open', 50.0, 0.0, 0.0, 0.0),
                  ('m1', 2, 'Closed', 10.0, 0.0, 0.0, 1.0), ('m2', 2, 'Closed', 10.0, 0.0, 0.0, 1.0),
                  ('n1', 3, 'Closed', 10.0, 0.0, 0.0, 1.0)])
    epc = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 30.0), ('X1', 1, 'Closed', 100.0, 0.0, 0.0, 10.0),
                 ('m1', 2, 'Closed', 30.0, 0.0, 0.0, 3.0), ('m2', 2, 'Open', 99.0, 0.0, 0.0, 9.0)])
    df, report = area.load_targets(epc, case)
    assert 'lo' not in df.index                                              # Open loads aren't scaled.
    # l1 starts from its EPC value (200), l2 from its seed (100): together they already make the 300 MW target.
    assert df['SMW_Target'].to_dict() == {'l1': 200.0, 'l2': 100.0, 'm1': 30.0, 'm2': 0.0, 'n1': 0.0}
    assert df.loc[['l1', 'l2'], 'SMvar_Target'].tolist() == [30.0, 10.0]
    assert (df[[f'{c}_Target' for c in ['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']]] == 0).all().all()
    assert (df['DistStatus_Target'] == 'Open').all() and (df['Status_Target'] == df['Status']).all()
    assert df['Method'].to_dict() == {'l1': 'Key', 'l2': 'Seed', 'm1': 'Key', 'm2': 'Key', 'n1': 'Seed'}
    r = report.set_index('AreaNumber')
    assert r.loc[1, 'Target MW'] == 300 and r.loc[1, 'Seed MW'] == 200 and r.loc[1, 'EPC-only MW'] == 100
    assert r.loc[2, 'Key Loads'] == 2 and r.loc[3, 'Target MW'] == 0
    assert (r['Unmet MW'].abs() < 1e-9).all()

def test_epc_loads_count_smw_and_smvar_only():
    # EPC l1 is S 100 + I 20 + Z 30 (Mvar 10/2/3): only the 100 MW / 10 Mvar count, in the area target and in the load's target.
    case = loads([('l1', 1, 'Closed', 80.0, 0.0, 0.0, 8.0)])
    epc = loads([('l1', 1, 'Closed', 100.0, 20.0, 30.0, 10.0)]).assign(IMvar=2.0, ZMvar=3.0)
    df, report = area.load_targets(epc, case)
    r = report.set_index('AreaNumber')
    assert r.loc[1, 'Target MW'] == 100 and r.loc[1, 'Target Mvar'] == 10
    assert df.loc['l1', 'SMW_Target'] == 100 and df.loc['l1', 'SMvar_Target'] == 10
    assert 'Seed Dist MW' not in report.columns

def test_a_blank_epc_smw_counts_as_0():
    case = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 10.0), ('l2', 1, 'Closed', 100.0, 0.0, 0.0, 10.0)])
    epc = loads([('l1', 1, 'Closed', np.nan, 0.0, 0.0, np.nan), ('l2', 1, 'Closed', 80.0, 0.0, 0.0, 8.0)])
    targets, _ = area.build_targets({'Load': epc, 'Gen': gens([])}, {'Load': case, 'Gen': gens([])})
    assert targets['Load']['SMW_Target'].notna().all() and targets['Load']['SMW_Target'].sum() == 80.0

def test_nonconforming_basis_is_the_epc_smw_only():
    case = loads([('s1', 1, 'Closed', 7.0, 0.0, 0.0, 1.0)])
    epc = loads([('s1', 1, 'Closed', 100.0, 20.0, 30.0, 10.0)]).assign(IMvar=2.0, ZMvar=3.0)
    df = area.nonconforming_targets(epc, case)
    assert df.loc['s1', 'Basis MW'] == 100 and df.loc['s1', 'Basis Mvar'] == 10 and df.loc['s1', 'SMW_Target'] == 100
    assert (df.loc['s1', ['IMW_Target', 'ZMW_Target', 'IMvar_Target', 'ZMvar_Target']] == 0).all()

def test_topology_targets_take_the_epc_smw_and_zero_its_i_and_z():
    common = dict(BusName='A', NomkV=230.0)
    load = lambda **kw: pd.DataFrame({'ObjectID': ['l1'], 'BusNum': 1, 'ID': '1', 'Status': 'Closed', 'SMW': 50.0, 'SMvar': 5.0, 'IMW': 0.0, 'IMvar': 0.0,
                                      'ZMW': 0.0, 'ZMvar': 0.0, 'DistStatus': 'Closed', 'DistMWInput': 3.0, 'DistMvarInput': 1.0, **common, **kw})
    gen = pd.DataFrame(columns=['ObjectID', 'BusNum', 'ID', 'Status', 'MWSetPoint', 'BusName', 'NomkV'])
    ld = topo.build_targets({'Gen': gen, 'Load': load(SMW=100.0, SMvar=10.0, IMW=20.0, IMvar=2.0, ZMW=30.0, ZMvar=3.0)}, {'Gen': gen, 'Load': load()})['Load']
    assert ld.loc['l1', 'SMW_Target'] == 100 and ld.loc['l1', 'SMvar_Target'] == 10
    assert (ld.loc['l1', ['IMW_Target', 'IMvar_Target', 'ZMW_Target', 'ZMvar_Target', 'DistMWInput_Target']] == 0).all()

def test_load_targets_area_with_no_seed_load_is_unmet():
    case = loads([('l1', 1, 'Open', 10.0, 0.0, 0.0, 0.0)])
    epc = loads([('l1', 1, 'Closed', 40.0, 0.0, 0.0, 0.0)])
    df, report = area.load_targets(epc, case)
    assert len(df) == 0 and report.set_index('AreaNumber').loc[1, 'Unmet MW'] == 40

def test_gen_targets():
    case = gens([('h1', 1, '1', 'Closed', 'HY', 0.0, 100.0, 0.0), ('h2', 1, '2', 'Closed', 'HY', 0.0, 300.0, 0.0),
                 ('w1', 1, '3', 'Closed', 'WT', 0.0, 50.0, 0.0), ('w2', 1, '4', 'Closed', 'WT', 0.0, 50.0, 0.0),
                 ('b1', 1, '5', 'Closed', 'BA', 0.0, 20.0, -20.0), ('b2', 1, '6', 'Closed', 'BA', 0.0, 60.0, -60.0),
                 ('t1', 1, '7', 'Closed', 'GT', 80.0, 100.0, 20.0), ('to', 1, '8', 'Open', 'GT', 0.0, 100.0, 0.0),
                 ('xx', 1, 'xx', 'Closed', 'BA', 500.0, 20000.0, -20000.0),
                 ('a', 2, '1', 'Closed', 'GT', 10.0, 100.0, 0.0), ('b', 2, '2', 'Closed', 'GT', 10.0, 100.0, 0.0)])
    epc = gens([('e1', 1, '1', 'Closed', 'HY', 200.0, 500.0, 0.0), ('e2', 1, '2', 'Closed', 'W3', 150.0, 200.0, 0.0),
                ('e3', 1, '3', 'Closed', 'ES', -40.0, 80.0, -80.0), ('e4', 1, '4', 'Closed', 'PV', 25.0, 30.0, 0.0),
                ('e5', 1, '5', 'Open', 'GT', 90.0, 100.0, 0.0),
                ('a', 2, '1', 'Closed', 'GT', 70.0, 100.0, 0.0), ('b', 2, '2', 'Open', 'GT', 50.0, 100.0, 0.0)])
    with saved_settings():
        area.MAKEUP = False
        df, report = area.gen_targets(epc, case)
    assert 'xx' not in df.index and 'to' not in df.index                    # Swing, and gens offline before and after.
    # Wind, storage: capacity types, all on at the same %. Hydro (a match type) with nothing to scale: committed like
    # Thermal (h2 alone covers 200 / 0.9). Thermal 0 MW: off. Matched area 2: EPC statuses and values.
    assert df['MWSetPoint_Target'].to_dict() == {'h1': 0.0, 'h2': 200.0, 'w1': 50.0, 'w2': 50.0, 'b1': -10.0, 'b2': -30.0,
                                                 't1': 0.0, 'a': 70.0, 'b': 0.0}
    assert df.loc[df['Status_Target'] == 'Open'].index.tolist() == ['h1', 't1', 'b']
    r = report.set_index(['AreaNumber', 'GenType'])['Unmet MW']
    assert r[(1, 'Wind')] == 50 and r[(1, 'Solar')] == 25 and r[(1, 'Hydro')] == 0 and r[(1, 'Thermal')] == 0
    # With make-up: Wind's 50 MW (capped) and Solar's 25 MW (no PWB units) go onto area 1's online Hydro (h2).
    # Storage (charging) takes none; area 2 is untouched. Unmet still says what each bucket's own units couldn't take.
    df, report = area.gen_targets(epc, case)
    assert df['MWSetPoint_Target'].to_dict() == {'h1': 0.0, 'h2': 275.0, 'w1': 50.0, 'w2': 50.0, 'b1': -10.0, 'b2': -30.0,
                                                 't1': 0.0, 'a': 70.0, 'b': 0.0}
    r = report.set_index(['AreaNumber', 'GenType'])
    assert r.loc[(1, 'Wind'), 'Unmet MW'] == 50 and r.loc[(1, 'Hydro'), 'Unmet MW'] == 0 and r.loc[(1, 'Hydro'), 'Make-up MW'] == 75
    assert r.loc[(1, 'Hydro'), 'Planned MW'] == 275 and (df['Planned MW'] == df['MWSetPoint_Target']).all()

def test_plan_makeup_commits_offline_units_and_follows_storage_down():
    # Area 1: Thermal target 100 (t1 online, MWMax 120; t2 offline, MWMax 500), 300 MW of EPC Solar with no PWB solar.
    # t1 takes 20 MW online; the other 280 MW re-dispatch Thermal at 400 MW, which brings t2 online.
    case = gens([('t1', 1, '1', 'Closed', 'GT', 100.0, 120.0, 0.0), ('t2', 1, '2', 'Open', 'GT', 0.0, 500.0, 0.0)])
    epc = gens([('t1', 1, '1', 'Closed', 'GT', 100.0, 120.0, 0.0), ('t2', 1, '2', 'Open', 'GT', 0.0, 500.0, 0.0),
                ('s1', 1, '3', 'Closed', 'PV', 300.0, 300.0, 0.0)])
    df, report = area.gen_targets(epc, case)
    assert abs(df['MWSetPoint_Target'].sum() - 400.0) < 1e-6 and df.loc['t2', 'Status_Target'] == 'Closed'
    r = report.set_index(['AreaNumber', 'GenType'])
    assert r.loc[(1, 'Solar'), 'Unmet MW'] == 300 and abs(r.loc[(1, 'Thermal'), 'Make-up MW'] - 300.0) < 1e-6
    # Storage charges -300 MW in the EPC (another key) but the PWB unit can take -100: the area's net gen keeps its total, so
    # Thermal (500) and Hydro (100) come down 200 MW between them, by MW.
    case = gens([('b1', 1, '1', 'Closed', 'BA', 0.0, 100.0, -100.0), ('t1', 1, '2', 'Closed', 'GT', 500.0, 1000.0, 0.0),
                 ('h1', 1, '3', 'Closed', 'HY', 100.0, 1000.0, 0.0)])
    epc = gens([('e1', 1, '9', 'Closed', 'BA', -300.0, 300.0, -300.0), ('t1', 1, '2', 'Closed', 'GT', 500.0, 1000.0, 0.0),
                ('h1', 1, '3', 'Closed', 'HY', 100.0, 1000.0, 0.0)])
    df, report = area.gen_targets(epc, case)
    assert df['MWSetPoint_Target'].round(6).to_dict() == {'b1': -100.0, 't1': round(500 - 200 * 5 / 6, 6), 'h1': round(100 - 200 / 6, 6)}
    assert abs(df['MWSetPoint_Target'].sum() - 300.0) < 1e-6                     # The EPC's net: -300 + 500 + 100.

def test_load_targets_scale_positive_and_negative_separately():
    # Mvar +100 / -98 (net +2) against +300 / -50: one net factor would blow up; split by sign it doesn't.
    case = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 100.0), ('l2', 1, 'Closed', 50.0, 0.0, 0.0, -98.0),
                  ('dc', 1, 'Closed', -500.0, 0.0, 0.0, 0.0)])
    # The negative EPC load is the PWB's dc load (same key), so it isn't netted as EPC-only.
    epc = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 300.0), ('dc', 1, 'Closed', -400.0, 0.0, 0.0, -50.0)])
    df, report = area.load_targets(epc, case)
    assert df['SMW_Target'].to_dict() == {'l1': 200.0, 'l2': 100.0, 'dc': -400.0}
    # Mvar -: dc starts from its EPC -50, l2 from its seed -98; together they scale to the EPC's -50.
    assert {k: round(v, 6) for k, v in df['SMvar_Target'].items()} == {'l1': 300.0, 'l2': round(-50 * 98 / 148, 6), 'dc': round(-50 * 50 / 148, 6)}
    r = report.iloc[0]
    assert (r['Target MW +'], r['Target MW -'], r['Seed MW +'], r['Seed MW -']) == (300, -400, 150, -500)
    assert all(abs(r[c]) < 1e-9 for c in report.columns if c.startswith('Unmet'))

def test_epc_only_negative_loads_are_netted_into_the_positive_load():
    # X1 (-30 MW, -5 Mvar) is only in the EPC: it lowers the area's positive load instead of going to the swing.
    case = loads([('l1', 1, 'Closed', 80.0, 0.0, 0.0, 10.0), ('n1', 1, 'Closed', -20.0, 0.0, 0.0, 0.0)])
    epc = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 20.0), ('n1', 1, 'Closed', -10.0, 0.0, 0.0, 0.0),
                 ('X1', 1, 'Closed', -30.0, 0.0, 0.0, -5.0)])
    df, report = area.load_targets(epc, case)
    assert df['SMW_Target'].to_dict() == {'l1': 70.0, 'n1': -10.0}      # n1 is in both cases: matched as a negative load.
    assert df.loc['l1', 'SMvar_Target'] == 15.0
    r = report.iloc[0]
    assert (r['Target MW'], r['Target MW +'], r['Target MW -'], r['EPC-only Netted MW']) == (60, 70, -10, -30)
    assert all(abs(r[c]) < 1e-9 for c in report.columns if c.startswith('Unmet'))

def test_load_targets_mvar_spread_by_mw_when_no_positive_mvar():
    case = loads([('l1', 1, 'Closed', 10.0, 0.0, 0.0, 0.0), ('l2', 1, 'Closed', 20.0, 0.0, 0.0, 0.0)])
    epc = loads([('e1', 1, 'Closed', 30.0, 0.0, 0.0, 30.0)])
    df, _ = area.load_targets(epc, case)
    assert df['SMvar_Target'].tolist() == [10.0, 20.0]

def test_load_targets_negative_target_with_no_negative_seed_load_is_unmet():
    # n1 is in both cases (not EPC-only), negative in the EPC but offline in the PWB: nothing negative to place it on.
    case = loads([('l1', 1, 'Closed', 10.0, 0.0, 0.0, 0.0), ('n1', 1, 'Open', 0.0, 0.0, 0.0, 0.0)])
    epc = loads([('l1', 1, 'Closed', 20.0, 0.0, 0.0, 0.0), ('n1', 1, 'Closed', -30.0, 0.0, 0.0, 0.0)])
    df, report = area.load_targets(epc, case)
    assert df['SMW_Target'].tolist() == [20.0]
    assert report['Unmet MW -'].tolist() == [-30.0] and report['Unmet MW +'].tolist() == [0.0]

def test_gen_targets_leave_dc_ties_alone():
    case = gens([('d1', 1, '1', 'Closed', 'DC', 300.0, 500.0, -500.0), ('g1', 1, '2', 'Closed', 'GT', 0.0, 100.0, 0.0)])
    epc = gens([('d1', 1, '1', 'Closed', 'DC', -400.0, 500.0, -500.0), ('g2', 1, '2', 'Closed', 'GT', 60.0, 100.0, 0.0)])
    df, report = area.gen_targets(epc, case)
    assert df.index.tolist() == ['g1'] and df.loc['g1', 'MWSetPoint_Target'] == 60
    assert report['GenType'].tolist() == ['Thermal']

def test_dist_gen_cleanup_clears_every_load_still_carrying_it():
    # lo is Open: cleared, but it injected nothing. ln is non-conforming and Closed: cleared, and its 5 MW leave the
    # case. lx's dist gen is already Open & 0: nothing to write. lz's is Open but not 0: cleared, 0 MW removed.
    case = loads([('lo', 1, 'Open', 50.0, 0.0, 0.0, 0.0), ('ln', 1, 'Closed', 50.0, 0.0, 0.0, 0.0),
                  ('lx', 1, 'Open', 50.0, 0.0, 0.0, 0.0), ('lz', 1, 'Closed', 50.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO', 'YES', 'YES'])
    case.loc[2, ['DistStatus', 'DistMWInput', 'DistMvarInput']] = ['Open', 0.0, 0.0]
    case.loc[3, 'DistStatus'] = 'Open'
    cleanup = area.dist_gen_cleanup(case)
    assert cleanup['ObjectID'].tolist() == ['lo', 'ln', 'lz'] and (cleanup['DistMWInput'] == 0).all() and (cleanup['DistStatus'] == 'Open').all()
    assert area.dist_gen_removed(cleanup) == 5.0
    assert area.dist_gen_removed(area.dist_gen_cleanup(case.iloc[2:3])) == 0.0  # Nothing to clear.

def test_unconverted_loads_warn_once_with_count_and_mw():
    case = loads([('a', 1, 'Closed', 50.0, 7.0, 0.0, 0.0), ('b', 1, 'Closed', 50.0, 0.0, 0.0, 0.0), ('c', 1, 'Open', 50.0, 9.0, 0.0, 0.0)])
    case.loc[1, ['DistStatus', 'DistMWInput', 'DistMvarInput']] = ['Open', 0.0, 0.0]
    case.loc[2, ['DistStatus', 'DistMWInput', 'DistMvarInput']] = ['Open', 0.0, 0.0]
    out = io.StringIO()
    with contextlib.redirect_stdout(out):
        n = area.warn_unconverted_loads(case)
    lines = out.getvalue().splitlines()
    assert n == 1 and len(lines) == 1 and lines[0].startswith('WARNING: Prepared.pwb has 1 Closed loads')
    assert '7 MW of I+Z' in lines[0] and 'Rerun step 01' in lines[0]

def test_converted_loads_and_open_dist_gen_print_no_unconverted_warning():
    case = loads([('a', 1, 'Closed', 50.0, 0.0, 0.0, 5.0), ('b', 1, 'Open', 50.0, 0.0, 0.0, 0.0)])
    case.loc[0, ['DistStatus', 'DistMWInput', 'DistMvarInput']] = ['Open', 0.0, 0.0]
    out = io.StringIO()
    with contextlib.redirect_stdout(out):
        n = area.warn_unconverted_loads(case)                 # b is Open with dist gen: it injects nothing.
    assert n == 0 and out.getvalue() == ''

def test_area_targets_through_the_engine():
    case = loads([('l1', 1, 'Closed', 60.0, 20.0, 20.0, 10.0), ('l2', 1, 'Closed', 100.0, 0.0, 0.0, 10.0)])
    epc = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 40.0)])
    case_gen = gens([('h1', 1, '1', 'Closed', 'WT', 0.0, 100.0, 0.0), ('h2', 1, '2', 'Closed', 'WT', 0.0, 300.0, 0.0)])
    epc_gen = gens([('e1', 1, '1', 'Closed', 'WT', 200.0, 500.0, 0.0)])
    targets, _ = area.build_targets({'Load': epc, 'Gen': epc_gen}, {'Load': case, 'Gen': case_gen})
    sa = FakeSimAuto()
    log = scale.iterate_to_gen_load_targets(sa, targets)
    assert results(log) == ['Solved'] * 3
    load, gen = targets['Load'], targets['Gen']
    assert load['SMW'].tolist() == [150.0, 150.0] and load['SMvar'].tolist() == [20.0, 20.0]
    assert (load[['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']] == 0).all().all()
    assert (load['DistStatus'] == 'Open').all() and (load['Status'] == 'Closed').all()
    assert gen['MWSetPoint'].tolist() == [50.0, 150.0] and (gen['Status'] == 'Closed').all()
    assert sa.applied[('Load', '1')]['DistStatus'] == 'Open'

def thermal_case():
    return gens([('g1', 1, '1', 'Closed', 'GT', 50.0, 100.0, 0.0), ('g2', 1, '2', 'Open', 'GT', 0.0, 400.0, 0.0),
                 ('g3', 1, '3', 'Open', 'GT', 0.0, 200.0, 0.0), ('g4', 1, '4', 'Open', 'GT', 0.0, 50.0, 0.0),
                 ('sc', 1, '5', 'Closed', 'SC', 0.0, 0.0, 0.0)])

def test_thermal_scaling_up_brings_biggest_units_on_first():
    # Nothing to scale (g1 online at 0 MW): the dispatch rules. Target 250 MW with 100 MW online, so the biggest (g2) comes on.
    case = thermal_case()
    case.loc[case['ObjectID'] == 'g1', 'MWSetPoint'] = 0.0
    df, report = area.gen_targets(gens([('e1', 1, '9', 'Closed', 'CC', 250.0, 300.0, 0.0)]), case)
    assert df['Status_Target'].to_dict() == {'g1': 'Closed', 'g2': 'Closed', 'sc': 'Closed'}
    r = report.set_index(['AreaNumber', 'GenType'])
    assert r.loc[(1, 'Thermal'), ['Committed', 'Decommitted', 'Unmet MW', 'Method']].tolist() == [1, 0, 0, 'Capacity']

def test_key_basis_scales_the_whole_group():
    # g3 is in the EPC (online, 150 MW) and e1 isn't in the PWB (100 MW): target 250. g3 starts from its EPC 150 MW and
    # status, g1 from its seed 50 MW; both scale by 250 / 200.
    epc = gens([('g3', 1, '3', 'Closed', 'GT', 150.0, 200.0, 0.0), ('e1', 1, '9', 'Closed', 'CC', 100.0, 300.0, 0.0)])
    df, report = area.gen_targets(epc, thermal_case())
    assert df['Status_Target'].to_dict() == {'g1': 'Closed', 'g3': 'Closed', 'sc': 'Closed'}
    assert df['MWSetPoint_Target'].to_dict() == {'g1': 62.5, 'g3': 187.5, 'sc': 0.0}
    assert df['Method'].to_dict() == {'g1': 'Seed', 'g3': 'Key', 'sc': 'Seed'}
    r = report.set_index(['AreaNumber', 'GenType'])
    assert r.loc[(1, 'Thermal'), ['Method', 'Basis MW', 'Key Gens', 'Unmet MW']].tolist() == ['Scaled', 200.0, 1, 0.0]

def test_group_with_no_key_scales_from_seed():
    # Decided by Fable 2026-11-15: no unit's key is in the EPC, so the seed dispatch (60 / 40) is scaled to the 50 MW target.
    case = gens([('g1', 1, '1', 'Closed', 'GT', 60.0, 100.0, 0.0), ('g2', 1, '2', 'Closed', 'GT', 40.0, 100.0, 0.0),
                 ('g3', 1, '3', 'Open', 'GT', 0.0, 500.0, 0.0)])
    df, report = area.gen_targets(gens([('e1', 1, '9', 'Closed', 'GT', 50.0, 300.0, 0.0)]), case)
    assert df['MWSetPoint_Target'].to_dict() == {'g1': 30.0, 'g2': 20.0} and (df['Status_Target'] == 'Closed').all()
    assert report['Method'].tolist() == ['Scaled']

def test_key_basis_at_mwmax_commits_more_units():
    # Target 400 from a basis of 200: g1 and g3 reach MWMax (100 + 200); the last 100 MW brings on the biggest offline unit.
    epc = gens([('g3', 1, '3', 'Closed', 'GT', 150.0, 200.0, 0.0), ('e1', 1, '9', 'Closed', 'CC', 250.0, 300.0, 0.0)])
    df, report = area.gen_targets(epc, thermal_case())
    assert df['MWSetPoint_Target'].to_dict() == {'g1': 100.0, 'g2': 100.0, 'g3': 200.0, 'sc': 0.0}
    assert df.loc['g2', 'Status_Target'] == 'Closed' and report['Unmet MW'].abs().max() < 1e-9

def test_key_basis_matched_unit_offline_in_epc():
    # g1 is Open in the EPC: with commit it goes offline, and g2 (seed) carries the target.
    case = gens([('g1', 1, '1', 'Closed', 'GT', 50.0, 100.0, 0.0), ('g2', 1, '2', 'Closed', 'GT', 50.0, 100.0, 0.0)])
    epc = gens([('g1', 1, '1', 'Open', 'GT', 0.0, 100.0, 0.0), ('e1', 1, '9', 'Closed', 'GT', 80.0, 100.0, 0.0)])
    df, _ = area.gen_targets(epc, case)
    assert df['Status_Target'].to_dict() == {'g1': 'Open', 'g2': 'Closed'} and df['MWSetPoint_Target'].tolist() == [0.0, 80.0]

def test_thermal_scaling_down_takes_smallest_units_off_first():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 100.0, 0.0), ('g2', 1, '2', 'Closed', 'ST', 200.0, 200.0, 0.0),
                 ('g3', 1, '3', 'Closed', 'CC', 300.0, 300.0, 0.0)])
    epc = gens([('e1', 1, '9', 'Closed', 'GT', 200.0, 300.0, 0.0)])
    df, _ = area.gen_targets(epc, case.assign(MWSetPoint=0.0))      # Online at 0 MW: nothing to scale, dispatch rules.
    assert df['Status_Target'].to_dict() == {'g1': 'Open', 'g2': 'Open', 'g3': 'Closed'}
    assert df['MWSetPoint_Target'].to_dict() == {'g1': 0.0, 'g2': 0.0, 'g3': 200.0}
    # With a key in the EPC (g1 at 100 MW), the group scales from its basis instead: all stay on, by 280 / 600.
    epc = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 100.0, 0.0), ('e1', 1, '9', 'Closed', 'GT', 180.0, 300.0, 0.0)])
    df, _ = area.gen_targets(epc, case)
    assert (df['Status_Target'] == 'Closed').all()
    assert df['MWSetPoint_Target'].round(6).to_dict() == {k: round(v * 280 / 600, 6) for k, v in {'g1': 100, 'g2': 200, 'g3': 300}.items()}

def test_commitment_off_keeps_seed_statuses():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 50.0, 100.0, 0.0), ('g2', 1, '2', 'Open', 'GT', 0.0, 400.0, 0.0)])
    epc = gens([('e1', 1, '9', 'Closed', 'GT', 150.0, 300.0, 0.0)])
    with saved_settings():
        area.COMMIT = False
        df, report = area.gen_targets(epc, case)
    assert df['Status_Target'].to_dict() == {'g1': 'Closed'} and df['MWSetPoint_Target'].tolist() == [100.0]
    assert report['Unmet MW'].tolist() == [50.0]

def test_commitment_headroom_and_capacity_types():
    case = gens([('t1', 1, '1', 'Closed', 'GT', 0.0, 100.0, 0.0), ('t2', 1, '2', 'Closed', 'GT', 0.0, 300.0, 0.0),
                 ('w1', 1, '3', 'Closed', 'WT', 0.0, 50.0, 0.0), ('w2', 1, '4', 'Open', 'WT', 0.0, 50.0, 0.0),
                 ('h1', 1, '5', 'Open', 'HY', 0.0, 100.0, 0.0), ('h2', 1, '6', 'Closed', 'HY', 0.0, 300.0, 0.0),
                 ('s1', 1, '7', 'Closed', 'PV', 3.0, 50.0, 0.0), ('s2', 1, '8', 'Open', 'PV', 0.0, 50.0, 0.0)])
    epc = gens([('e1', 1, '1', 'Closed', 'GT', 280.0, 300.0, 0.0), ('e2', 1, '2', 'Closed', 'WT', 10.0, 100.0, 0.0),
                ('e3', 1, '3', 'Closed', 'HY', 40.0, 100.0, 0.0)])
    df, _ = area.gen_targets(epc, case)
    st = df['Status_Target'].to_dict()
    assert st['t1'] == 'Closed' and st['t2'] == 'Closed'   # 280 MW is over 0.9 of the 300 MW unit: both stay on.
    assert st['w1'] == 'Closed' and st['w2'] == 'Closed'   # Wind: all units on, even for 10 MW.
    assert df.loc[['w1', 'w2'], 'MWSetPoint_Target'].tolist() == [5.0, 5.0]
    assert 'h1' not in st and df.loc['h2', 'MWSetPoint_Target'] == 40.0  # Hydro (a match type): committed like Thermal.
    assert st['s1'] == 'Open' and 's2' not in st           # Solar target 0: all off.
    assert df.loc['s1', 'MWSetPoint_Target'] == 0
    # Storage target 0: offline, like Wind and Solar. Left out of off_at_zero_types: seed statuses, at 0 MW.
    case = gens([('b1', 1, '1', 'Closed', 'BA', 5.0, 50.0, -50.0), ('b2', 1, '2', 'Open', 'BA', 0.0, 50.0, -50.0)])
    df, _ = area.gen_targets(gens([]), case)
    assert df['Status_Target'].to_dict() == {'b1': 'Open'} and df['MWSetPoint_Target'].tolist() == [0.0]
    with saved_settings():
        area.OFF_AT_ZERO_TYPES = {'Wind', 'Solar'}
        df, _ = area.gen_targets(gens([]), case)
    assert df['Status_Target'].to_dict() == {'b1': 'Closed'} and df['MWSetPoint_Target'].tolist() == [0.0]

def test_capacity_types_ignore_unit_epc_values_and_run_flat():
    # Wind: keyed units at very different EPC MW (w3 Open there) all come online at the area total's same % of MWMax.
    case = gens([('w1', 1, '1', 'Closed', 'WT', 90.0, 100.0, 0.0), ('w2', 1, '2', 'Closed', 'WT', 0.0, 200.0, 0.0),
                 ('w3', 1, '3', 'Open', 'WT', 0.0, 100.0, 0.0),
                 ('b1', 1, '4', 'Closed', 'BA', 10.0, 50.0, -50.0), ('b2', 1, '5', 'Closed', 'BA', -5.0, 50.0, -150.0)])
    epc = gens([('w1', 1, '1', 'Closed', 'WT', 80.0, 100.0, 0.0), ('w2', 1, '2', 'Closed', 'WT', 20.0, 200.0, 0.0),
                ('w3', 1, '3', 'Open', 'WT', 0.0, 100.0, 0.0),
                ('b1', 1, '4', 'Closed', 'BA', -50.0, 50.0, -50.0), ('b2', 1, '5', 'Closed', 'BA', 10.0, 50.0, -150.0)])
    with saved_settings():
        area.MAKEUP = False
        df, report = area.gen_targets(epc, case)
    assert (df['Status_Target'] == 'Closed').all()
    assert df.loc[['w1', 'w2', 'w3'], 'MWSetPoint_Target'].tolist() == [25.0, 50.0, 25.0]      # 100 MW: 25% of each.
    assert df.loc[['b1', 'b2'], 'MWSetPoint_Target'].tolist() == [-10.0, -30.0]                # Charging 40 MW: by |MWMin|.
    assert set(report.loc[report['GenType'].isin(['Wind', 'Battery Storage']), 'Method']) == {'Capacity'}
    # Not a capacity type: the same wind units match key by key instead, and w3 stays Open as in the EPC.
    with saved_settings():
        area.MAKEUP, area.CAPACITY_TYPES = False, {'Battery Storage'}
        df, _ = area.gen_targets(epc, case)
    assert df.loc[['w1', 'w2'], 'MWSetPoint_Target'].tolist() == [80.0, 20.0] and 'w3' not in df.index

def test_match_types_take_epc_values_key_by_key_and_fill_from_seed():
    # Thermal: t1 and t2 keyed (t2 Open in the EPC), t3 not in the EPC keeps its seed MW. Total 240 = 150 + 0 + 60 + 30
    # more: the basis (210) is scaled by 240 / 210.
    case = gens([('t1', 1, '1', 'Closed', 'GT', 100.0, 300.0, 0.0), ('t2', 1, '2', 'Closed', 'GT', 100.0, 300.0, 0.0),
                 ('t3', 1, '3', 'Closed', 'GT', 60.0, 300.0, 0.0)])
    epc = gens([('t1', 1, '1', 'Closed', 'GT', 150.0, 300.0, 0.0), ('t2', 1, '2', 'Open', 'GT', 80.0, 300.0, 0.0),
                ('e9', 1, '9', 'Closed', 'GT', 90.0, 300.0, 0.0)])
    df, _ = area.gen_targets(epc, case)
    assert df['Status_Target'].to_dict() == {'t1': 'Closed', 't2': 'Open', 't3': 'Closed'}
    assert df['MWSetPoint_Target'].round(9).to_dict() == {'t1': round(150 * 240 / 210, 9), 't2': 0.0, 't3': round(60 * 240 / 210, 9)}
    # Every key matched: each unit gets exactly its EPC value.
    epc = gens([('t1', 1, '1', 'Closed', 'GT', 150.0, 300.0, 0.0), ('t2', 1, '2', 'Open', 'GT', 80.0, 300.0, 0.0),
                ('t3', 1, '3', 'Closed', 'GT', 20.0, 300.0, 0.0)])
    df, _ = area.gen_targets(epc, case)
    assert df['MWSetPoint_Target'].to_dict() == {'t1': 150.0, 't2': 0.0, 't3': 20.0}

def test_capacity_share_leaves_out_units_outside_the_main_island():
    # w3 is on bus 3, island 2: excluded and turned off; w1 and w2 carry the whole 150 MW at the same %.
    case = gens([('w1', 1, '1', 'Closed', 'WT', 10.0, 100.0, 0.0), ('w2', 1, '2', 'Closed', 'WT', 10.0, 200.0, 0.0),
                 ('w3', 1, '3', 'Closed', 'WT', 10.0, 300.0, 0.0)]).assign(BusNum=[1, 2, 3])
    epc = gens([('w1', 1, '1', 'Closed', 'WT', 50.0, 100.0, 0.0), ('w2', 1, '2', 'Closed', 'WT', 50.0, 200.0, 0.0),
                ('w3', 1, '3', 'Closed', 'WT', 50.0, 300.0, 0.0)])
    targets, _ = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case},
                                    island_by_bus=pd.Series({1: 1, 2: 1, 3: 2}))
    g = targets['Gen']
    assert g.loc[['w1', 'w2'], 'MWSetPoint_Target'].tolist() == [50.0, 100.0]
    assert g.loc['w3', 'MWSetPoint_Target'] == 0 and g.loc['w3', 'Status_Target'] == 'Open'
    assert scale.is_island_reason(g.loc[['w3'], 'ExclusionReason']).all()

def test_nonconforming_loads_are_left_alone():
    # l2 / e2 are non-conforming (AGC NO): kept at seed value, out of both area totals, out of balancing.
    case = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 10.0), ('l2', 1, 'Closed', 500.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO '])
    epc = loads([('e1', 1, 'Closed', 150.0, 0.0, 0.0, 15.0), ('e2', 1, 'Closed', 900.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO'])
    targets, reports = area.build_targets({'Load': epc, 'Gen': gens([])}, {'Load': case, 'Gen': gens([])})
    assert targets['Load'].index.tolist() == ['l1'] and targets['Load'].loc['l1', 'SMW_Target'] == 150
    r = reports['area_load'].iloc[0]
    assert (r['Target MW'], r['Left As Is MW']) == (150, 500)
    bal = area.balance_targets({'Gen': gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 300.0, 0.0)]), 'Load': case}, 50.0, [])
    assert bal['Load'].empty                                                     # Balancing never moves load.
    with saved_settings():
        area.SCALE_NONCONFORMING = True
        targets, reports = area.build_targets({'Load': epc, 'Gen': gens([])}, {'Load': case, 'Gen': gens([])})
        assert targets['Load']['SMW_Target'].sum() == 1050 and reports['nonconforming'] == set() and not targets['Load']['Nonconforming'].any()

def test_build_targets_leaves_slack_bus_gens_out_of_both_cases():
    # Identical cases: the slack-bus unit's 300 MW is not loaded onto g1, and the area still matches by key.
    case = gens([('s1', 1, '1', 'Closed', 'GT', 300.0, 500.0, 0.0), ('g1', 1, '2', 'Closed', 'GT', 80.0, 100.0, 0.0)]).assign(BusNum=[5, 6])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': case}, {'Load': loads([]), 'Gen': case}, swing_buses=[5])
    assert targets['Gen'].index.tolist() == ['g1'] and targets['Gen'].loc['g1', 'MWSetPoint_Target'] == 80
    assert reports['area_gen']['Method'].tolist() == ['Scaled'] and reports['area_gen']['Unmet MW'].tolist() == [0.0]

def test_nonconforming_is_decided_by_key():
    # l2 is non-conforming in the PWB only. The EPC's l2 (900 MW) must not be loaded onto l1: it is out of the area's
    # target (100), and l2 itself is set to its EPC value 1:1.
    case = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 0.0), ('l2', 1, 'Closed', 500.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO'])
    epc = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 0.0), ('l2', 1, 'Closed', 900.0, 0.0, 0.0, 0.0)])
    targets, reports = area.build_targets({'Load': epc, 'Gen': gens([])}, {'Load': case, 'Gen': gens([])})
    t = targets['Load']
    assert t.index.tolist() == ['l1', 'l2'] and t.loc['l1', 'SMW_Target'] == 100 and t.loc['l2', 'SMW_Target'] == 900
    assert t['Method'].tolist() == ['Key', 'EPC 1:1'] and t['Nonconforming'].tolist() == [False, True]
    assert reports['nonconforming'] == {'l2'}
    r = reports['area_load'].iloc[0]
    assert (r['Target MW'], r['Planned MW']) == (100, 100)
    assert (r['Set To EPC MW'], r['Set To EPC Loads'], r['Left As Is MW']) == (900, 1, 0)

def test_blank_epc_codes_stop_the_run_before_scaling():
    case = gens([('w1', 1, '1', 'Closed', 'WT', 40.0, 50.0, 0.0), ('t1', 1, '2', 'Closed', 'GT', 80.0, 100.0, 0.0), ('h1', 1, '3', 'Closed', 'HY', 80.0, 100.0, 0.0)])
    epc = case.assign(UnitTypeCode='', ObjectID=['a', 'b', 'c'])
    try:
        area.gen_targets(epc, case)
        assert False
    except ValueError as e:
        assert '200 MW' in str(e) and 'unit_gentype' in str(e) and 'fuel_gentype' in str(e)
    # 40% of the MW has no type: under the limit.
    epc = epc.assign(UnitTypeCode=['', 'GT', 'HY'], MWSetPoint=[80.0, 60.0, 60.0])
    df, _ = area.gen_targets(epc, case)
    assert len(df) == 3
    with saved_settings():
        area.MAX_UNKNOWN_SHARE = 1.0   # Off.
        area.gen_targets(epc.assign(UnitTypeCode=''), case)

def test_target_zero_takes_every_committed_unit_offline():
    rng = np.random.default_rng(0)
    for _ in range(300):
        caps = np.round(rng.uniform(1, 900, rng.integers(2, 12)), rng.integers(0, 3))
        rows = gens([(f'g{i}', 1, str(i), 'Closed', 'GT', 1.0, c, 0.0) for i, c in enumerate(caps)])
        assert not area.commit(rows, 0.0, 'Thermal', set()).any(), caps.tolist()
    case = gens([('t1', 1, '1', 'Closed', 'GT', 10.0, 431.58, 0.0), ('t2', 1, '2', 'Closed', 'ST', 10.0, 556.86, 0.0), ('t3', 1, '3', 'Closed', 'CC', 10.0, 99.99, 0.0)])
    df, _ = area.gen_targets(gens([]), case)
    assert (df['Status_Target'] == 'Open').all()
    # A unit with no capacity is left alone.
    rows = gens([('b1', 1, '1', 'Closed', 'BA', 5.0, 50.0, -50.0), ('c1', 1, '2', 'Closed', 'SC', 0.0, 0.0, 0.0)])
    assert area.commit(rows, 0.0, 'Thermal', set()).tolist() == [False, True]

def test_decommit_skips_a_unit_which_cannot_go_and_tries_the_next():
    # Need 520 MW. A (500) is first in line (offline in the EPC) but removing it leaves 100. B (50) can go; C can't.
    rows = gens([('A', 1, '1', 'Closed', 'GT', 1.0, 500.0, 0.0), ('B', 1, '2', 'Closed', 'GT', 1.0, 50.0, 0.0), ('C', 1, '3', 'Closed', 'GT', 1.0, 50.0, 0.0)])
    assert area.commit(rows, 468.0, 'Thermal', {'B', 'C'}).tolist() == [True, False, True]

def test_epc_units_on_islanded_seed_buses_are_left_out_of_the_targets():
    # (a) EPC hydro e2 and EPC load e2 on bus 2 (island 2) are out of Target MW and the Note says so; (c) bus 4 has a blank
    # island: dropped; (b) EPC X1 on bus 77 (not in the seed) still counts, netted when negative. Not counted in the Note:
    # an Open EPC gen, a DC tie, a non-conforming load.
    case_gen = gens([('h1', 1, '1', 'Closed', 'HY', 100.0, 200.0, 0.0), ('h2', 1, '2', 'Closed', 'HY', 100.0, 200.0, 0.0)]).assign(BusNum=[1, 2])
    epc_gen = gens([('e1', 1, '1', 'Closed', 'HY', 150.0, 200.0, 0.0), ('e2', 1, '2', 'Closed', 'HY', 40.0, 200.0, 0.0),
                    ('e3', 1, '3', 'Closed', 'HY', 30.0, 200.0, 0.0), ('e4', 1, '4', 'Open', 'HY', 20.0, 200.0, 0.0),
                    ('e5', 1, '5', 'Closed', 'DC', 10.0, 200.0, -200.0)]).assign(BusNum=[1, 2, 4, 2, 2])
    case_load = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 10.0)])
    epc_load = loads([('l1', 1, 'Closed', 120.0, 0.0, 0.0, 10.0), ('e2', 1, 'Closed', 50.0, 0.0, 0.0, 5.0),
                      ('X1', 1, 'Closed', -30.0, 0.0, 0.0, 0.0), ('nc', 1, 'Closed', 7.0, 0.0, 0.0, 0.0)]).assign(BusNum=[1, 2, 77, 2])
    epc_load.loc[3, 'AGC'] = 'NO'
    islands = pd.Series([1.0, 2.0, np.nan], index=[1, 2, 4])
    with saved_settings():
        area.SCALE_NONCONFORMING = False
        targets, reports = area.build_targets({'Gen': epc_gen, 'Load': epc_load}, {'Gen': case_gen, 'Load': case_load}, island_by_bus=islands)
    g = reports['area_gen'].set_index('GenType')
    assert g.loc['Hydro', 'Target MW'] == 150.0                                  # e2 (40) and e3 (30, blank island) are out.
    assert reports['area_load'].set_index('AreaNumber').loc[1, 'Target MW'] == 90.0     # 120 - 30 netted; e2 (50) out; nc not counted.
    assert reports['islanded_epc'].to_dict() == {(1, 'Hydro'): 70.0, (1, 'Load'): 50.0}
    out = area.areas(None, reports, targets).set_index('Group')
    assert 'EPC outside main island, left out: 70.0 MW' in out.loc['Hydro', 'Note']
    assert 'EPC outside main island, left out: 50.0 MW' in out.loc['Load', 'Note']
    # (d) Without islands nothing is left out.
    targets, reports = area.build_targets({'Gen': epc_gen, 'Load': epc_load}, {'Gen': case_gen, 'Load': case_load})
    assert reports['islanded_epc'].empty and reports['area_gen'].set_index('GenType').loc['Hydro', 'Target MW'] == 220.0
    assert scale.epc_off_island(epc_load, None).sum() == 0

def test_islanded_units_take_no_share_of_the_area_targets():
    # Two hydro units at 100 MW, h2 in island 2. The EPC's 150 MW all goes to h1; h2 stays at 100, excluded.
    # Same for load: l2 islanded, l1 takes the whole 300 MW.
    case_gen = gens([('h1', 1, '1', 'Closed', 'HY', 100.0, 200.0, 0.0), ('h2', 1, '2', 'Closed', 'HY', 100.0, 200.0, 0.0),
                     ('dc', 1, '3', 'Closed', 'DC', 20.0, 50.0, -50.0)]).assign(BusNum=[1, 2, 2])
    epc_gen = gens([('e1', 1, '1', 'Closed', 'HY', 150.0, 200.0, 0.0)])
    case_load = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 10.0), ('l2', 1, 'Closed', 50.0, 0.0, 0.0, 5.0)]).assign(BusNum=[1, 3])
    epc_load = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)])
    islands = pd.Series([1.0, 2.0, 2.0], index=[1, 2, 3])
    targets, reports = area.build_targets({'Gen': epc_gen, 'Load': epc_load}, {'Gen': case_gen, 'Load': case_load}, island_by_bus=islands)
    g, l = targets['Gen'], targets['Load']
    assert g.loc['h1', 'MWSetPoint_Target'] == 150.0 and g.loc['h2', 'MWSetPoint_Target'] == 0.0 and g.loc['h2', 'Status_Target'] == 'Open'
    assert g.loc['h2', 'ExclusionReason'] == 'Island 2, not the main island' and g.loc['h1', 'ExclusionReason'] == ''
    assert l.loc['l1', 'SMW_Target'] == 300.0 and l.loc['l2', 'SMW_Target'] == 0.0 and l.loc['l2', 'ExclusionReason'] != ''
    assert reports['area_load'].set_index('AreaNumber').loc[1, 'Seed MW'] == 100.0                  # l2 left out of the seed total.
    assert g.loc['dc', 'Status_Target'] == 'Open'                                # A DC tie in a second island goes off too.
    assert reports['islanded'].to_dict() == {(1, 'Hydro'): 100.0, (1, 'Load'): 50.0, (1, 'Unknown'): 20.0}
    # Without islands, the two units share the target.
    targets, _ = area.build_targets({'Gen': epc_gen, 'Load': epc_load}, {'Gen': case_gen, 'Load': case_load})
    assert targets['Gen'].loc[['h1', 'h2'], 'MWSetPoint_Target'].tolist() == [75.0, 75.0]

def test_plan_makeup_commitment_leaves_offline_solar_and_wind_off():
    # Night: EPC Thermal 300 MW on a key the PWB lacks, EPC Solar and Wind at 0. Thermal is short (t1 MWMax 200, t2 offline
    # with 30 MW); s1 / w1 are online in a daytime seed. The leftover must not bring any Solar / Wind unit online.
    case = gens([('t1', 1, '1', 'Closed', 'GT', 150.0, 200.0, 0.0), ('t2', 1, '4', 'Open', 'GT', 0.0, 30.0, 0.0),
                 ('s1', 1, '2', 'Closed', 'PV', 100.0, 400.0, 0.0), ('w1', 1, '3', 'Closed', 'WT', 50.0, 400.0, 0.0)])
    epc = gens([('eT', 1, '9', 'Closed', 'GT', 300.0, 400.0, 0.0),
                ('s1', 1, '2', 'Closed', 'PV', 0.0, 400.0, 0.0), ('w1', 1, '3', 'Closed', 'WT', 0.0, 400.0, 0.0)])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case})
    g = targets['Gen']
    assert g.loc[['s1', 'w1'], 'Status_Target'].tolist() == ['Open', 'Open'] and g.loc[['s1', 'w1'], 'MWSetPoint_Target'].tolist() == [0.0, 0.0]
    assert g.loc[['t1', 't2'], 'MWSetPoint_Target'].tolist() == [200.0, 30.0]
    g['MWSetPoint'] = g['MWSetPoint_Target']                                      # As if scaling reached every target.
    final = case.set_index('ObjectID').loc[g.index].assign(Status=g['Status_Target'], MWSetPoint=g['MWSetPoint_Target']).reset_index()
    out = area.areas({'Gen': final, 'Load': loads([]).assign(NomkV=1.0)}, reports, targets).set_index('Group')
    assert out.loc['Thermal', 'Unmet MW'] == 70.0 and out.loc['Solar', 'Unmet MW'] == 0.0 and out.loc['Wind', 'Unmet MW'] == 0.0
    assert not out['Note'].str.contains('unexplained').any(), out['Note']

def test_hydro_is_a_match_type_off_at_zero_and_committed_by_plan_makeup():
    # Hydro target 0: its seed-online unit goes offline (match types commit like Thermal), not kept on at 0 MW.
    case = gens([('h1', 1, '1', 'Closed', 'HY', 20.0, 100.0, 0.0), ('t1', 1, '2', 'Closed', 'GT', 100.0, 200.0, 0.0)])
    df, _ = area.gen_targets(gens([('eT', 1, '9', 'Closed', 'GT', 100.0, 200.0, 0.0)]), case)
    assert df.loc['h1', 'Status_Target'] == 'Open' and df.loc['h1', 'MWSetPoint_Target'] == 0
    # Thermal 300 MW short of capacity (t1 MWMax 200): the plan make-up re-dispatches Hydro, bringing offline h2 online.
    case = gens([('t1', 1, '1', 'Closed', 'GT', 150.0, 200.0, 0.0), ('h2', 1, '2', 'Open', 'HY', 0.0, 300.0, 0.0)])
    df, report = area.gen_targets(gens([('eT', 1, '9', 'Closed', 'GT', 300.0, 400.0, 0.0)]), case)
    assert df.loc['h2', 'Status_Target'] == 'Closed' and df.loc['h2', 'MWSetPoint_Target'] == 100.0
    assert report.set_index('GenType').loc['Hydro', 'Make-up MW'] == 100.0

class LowBusSimAuto(SwingSimAuto):
    """SwingSimAuto with the buses in `low` at 0.85 pu (below the scaling band)."""
    low = frozenset()

    def GetParametersMultipleElementRect(self, table, fields, filter_group):
        err, rows = super().GetParametersMultipleElementRect(table, fields, filter_group)
        if table == 'Bus' and 'Vpu' in fields:
            n, v = fields.index('Number'), fields.index('Vpu')
            rows = [[('0.85' if i == v and int(r[n]) in self.low else x) for i, x in enumerate(r)] for r in rows]
        return err, rows

def nonc_hour(case_rows, epc_rows, flags=('YES', 'YES'), epc_flags=('YES', 'YES'), low=(), buses=(3, 4, 5, 6), seed_flags=None):
    """
    Area 1: conforming l1 (seed 200, EPC 120) and the loads of case_rows / epc_rows after it (loads() rows, area 1), on buses 3, 4, ...
    Scales a SwingSimAuto hour through the engine. Returns (sa, targets, reports, areas sheet).
    """
    base = [('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]
    dec = lambda df, agc: df.assign(BusNum=list(buses)[:len(df)], NomkV=230.0, BusName='b', AGC=agc)
    case = dec(loads(base + case_rows), ['YES'] + list(seed_flags or flags))
    epc = dec(loads([('l1', 1, 'Closed', 120.0, 0.0, 0.0, 12.0)] + epc_rows), ['YES'] + list(epc_flags))
    gen = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0)]).assign(BusNum=1, MW=0.0, NomkV=230.0, BusName='b')
    sa = LowBusSimAuto(gen, case)
    sa.low = frozenset(low)
    targets, reports = area.build_targets({'Load': epc, 'Gen': gen}, {'Load': case, 'Gen': gen})
    with contextlib.redirect_stdout(io.StringIO()):
        scale.iterate_to_gen_load_targets(sa, targets)
        mk_runs = area.makeup(sa, targets, nonconforming=reports['nonconforming'])
        sheet = area.areas(area.read_case(sa), reports, targets, mk_runs)
    return sa, targets, reports, sheet, [r.moved for r in mk_runs]

def test_load_flagged_only_in_the_epc_is_non_conforming_and_out_of_the_ledger():
    # ss is conforming in the PWB, flagged in the EPC (80 MW): out of the area's target and of Final MW, and set to 80 MW 1:1.
    sa, targets, reports, sheet, _ = nonc_hour([('ss', 1, 'Closed', 50.0, 0.0, 0.0, 0.0)], [('ss', 1, 'Closed', 80.0, 0.0, 0.0, 0.0)],
                                               flags=('YES',), epc_flags=('NO',))
    assert reports['nonconforming'] == {'ss'} and targets['Load'].loc['ss', 'Method'] == 'EPC 1:1'
    assert sa.table('Load').set_index('BusNum').loc[4, 'SMW'] == 80.0 and sa.table('Load').set_index('BusNum').loc[3, 'SMW'] == 120.0
    row = sheet[sheet['Group'] == 'Load'].iloc[0]
    assert (row['Target MW'], row['Final MW'], row['Error MW'], row['Online Final']) == (120.0, 120.0, 0.0, 1.0)
    assert 'unexplained' not in row['Note'] and 'non-conforming set to EPC: 80.0 MW (1 load)' in row['Note']
    assert not sheet['Note'].str.contains('unexplained').any()

def test_an_area_with_only_nonconforming_loads_gets_a_load_row_for_the_note():
    case = loads([('k', 2, 'Closed', 40.0, 0.0, 0.0, 0.0)]).assign(AGC='NO')
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': gens([])}, {'Load': case, 'Gen': gens([])})
    out = area.areas(None, reports, targets)
    assert out['Group'].tolist() == ['Load'] and out['AreaNumber'].tolist() == [2]
    assert out['Note'].iloc[0] == 'non-conforming left as is: 40.0 MW'
    assert len(targets['Load']) == 0

def test_nonconforming_load_held_for_voltage_keeps_the_ledger_and_the_makeup_off_conforming_loads():
    # ss (bus 4) sags below the voltage band: held, then excluded 30 MW short of its EPC value. The make-up must not put
    # those 30 MW on l1, and the ledger (which leaves ss out) still holds.
    sa, targets, reports, sheet, mk = nonc_hour([('ss', 1, 'Closed', 50.0, 0.0, 0.0, 0.0)], [('ss', 1, 'Closed', 80.0, 0.0, 0.0, 0.0)],
                                                flags=('NO',), epc_flags=('NO',), low=(4,))
    t = targets['Load'].loc['ss']
    assert t['ExclusionReason'] != '' and t['SMW'] == 50.0 and t['SMW_Target'] == 80.0
    assert all(len(m['Load']) == 0 for m in mk) and sa.table('Load').set_index('BusNum').loc[3, 'SMW'] == 120.0
    row = sheet[sheet['Group'] == 'Load'].iloc[0]
    assert (row['Target MW'], row['Not Reached MW'], row['Make-up MW'], row['Final MW']) == (120.0, 0.0, 0.0, 120.0)
    assert not sheet['Note'].str.contains('unexplained').any()
    # The make-up's own shortfall list leaves it out directly, too.
    final = area.read_case(sa)
    assert area.makeup_targets(final, targets, swing_buses=[], nonconforming=reports['nonconforming']) is None

def test_makeup_leaves_non_conforming_loads_alone_unless_scale_nonconforming():
    # l1 (bus 3) is held 80 MW short. l2 is conforming; ss (flagged in the EPC only) has headroom: it must not take a share.
    rows = [('l2', 1, 'Closed', 100.0, 0.0, 0.0, 0.0), ('ss', 1, 'Closed', 50.0, 0.0, 0.0, 0.0)]
    epc_rows = [('l2', 1, 'Closed', 100.0, 0.0, 0.0, 0.0), ('ss', 1, 'Closed', 50.0, 0.0, 0.0, 0.0)]
    with saved_settings():
        sa, targets, reports, sheet, mk = nonc_hour(rows, epc_rows, flags=('YES', 'YES'), epc_flags=('YES', 'NO'), low=(3,))
    assert reports['nonconforming'] == {'ss'} and targets['Load'].loc['l1', 'ExclusionReason'] != ''
    assert [m['Load'].index.tolist() for m in mk] == [['l2']] and abs(mk[0]['Load'].loc['l2', 'SMW_Target'] - 20.0) < 1e-9
    out = sa.table('Load').set_index('BusNum')
    assert abs(out.loc[4, 'SMW'] - 20.0) < 1e-9 and out.loc[5, 'SMW'] == 50.0
    assert not sheet['Note'].str.contains('unexplained').any(), sheet['Note']
    # With scale_nonconforming they are ordinary loads again: ss takes a share of the make-up too.
    with saved_settings():
        area.SCALE_NONCONFORMING = True
        sa, targets, reports, sheet, mk = nonc_hour(rows, epc_rows, flags=('YES', 'YES'), epc_flags=('YES', 'NO'), low=(3,))
    assert reports['nonconforming'] == set() and sorted(mk[0]['Load'].index) == ['l2', 'ss']
    assert not sheet['Note'].str.contains('unexplained').any(), sheet['Note']

def test_non_conforming_loads_open_in_both_cases_are_not_target_rows():
    # s1 is Open in the seed and in the EPC: nothing to change, so no target row (no write, no exception, not counted).
    # s2 is Closed in the EPC: the one row, so the Note says '1 load'.
    sa, targets, reports, sheet, _ = nonc_hour(
        [('s1', 1, 'Open', 7.0, 0.0, 0.0, 0.0), ('s2', 1, 'Open', 7.0, 0.0, 0.0, 0.0)],
        [('s1', 1, 'Open', 9.0, 0.0, 0.0, 0.0), ('s2', 1, 'Closed', 25.0, 0.0, 0.0, 0.0)], flags=('NO', 'NO'), epc_flags=('NO', 'NO'), low=(4,))
    assert reports['nonconforming'] == {'s1', 's2'} and 's1' not in targets['Load'].index
    assert targets['Load'].loc['s2', 'Status_Target'] == 'Closed'
    row = sheet[sheet['Group'] == 'Load'].iloc[0]
    assert 'non-conforming set to EPC: 25.0 MW (1 load)' in row['Note']

def test_nonconforming_loads_take_the_epc_status_and_those_without_a_key_are_left_alone():
    # s1: Open in the seed, Closed in the EPC (SMW 30; its I 5 and Z 5 don't count): closed in. s2: Closed in the seed, Open in the EPC: opened.
    # s3: no EPC key: untouched. Dist gen of the 1:1 loads goes Open & 0.
    sa, targets, reports, sheet, _ = nonc_hour(
        [('s1', 1, 'Open', 7.0, 0.0, 0.0, 0.0), ('s2', 1, 'Closed', 60.0, 0.0, 0.0, 0.0), ('s3', 1, 'Closed', 30.0, 0.0, 0.0, 0.0)],
        [('s1', 1, 'Closed', 30.0, 5.0, 5.0, 4.0), ('s2', 1, 'Open', 99.0, 0.0, 0.0, 0.0)],
        flags=('NO', 'NO', 'NO'), epc_flags=('NO', 'NO'))
    t = targets['Load']
    assert t.loc[['s1', 's2'], 'Status_Target'].tolist() == ['Closed', 'Open'] and 's3' not in t.index
    assert t.loc[['s1', 's2'], 'SMW_Target'].tolist() == [30.0, 0.0] and t.loc['s1', 'SMvar_Target'] == 4.0
    assert (t.loc[['s1', 's2'], ['IMW_Target', 'ZMW_Target', 'DistMWInput_Target']] == 0).all().all()
    assert t.loc[['s1', 's2'], 'DistStatus_Target'].tolist() == ['Open', 'Open']
    out = sa.table('Load').set_index('BusNum')
    assert (out.loc[4, 'Status'], out.loc[4, 'SMW'], out.loc[4, 'SMvar'], out.loc[4, 'DistStatus']) == ('Closed', 30.0, 4.0, 'Open')
    assert (out.loc[5, 'Status'], out.loc[5, 'SMW']) == ('Open', 0.0)
    assert (out.loc[6, 'Status'], out.loc[6, 'SMW']) == ('Closed', 30.0)
    row = sheet[sheet['Group'] == 'Load'].iloc[0]
    assert 'non-conforming set to EPC: 30.0 MW (2 loads)' in row['Note'] and 'non-conforming left as is: 30.0 MW' in row['Note']
    assert row['Final MW'] == row['Target MW'] == 120.0 and not sheet['Note'].str.contains('unexplained').any()

def test_loss_estimate_counts_one_to_one_loads_at_their_epc_value_and_others_at_seed():
    # Area 1 seed load: l1 200 + ss 50 (EPC 100) + kl 30 (no key, left as is) = 280; target: l1 120 + ss 100 + kl 30 = 250.
    case = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 0.0), ('ss', 1, 'Closed', 50.0, 0.0, 0.0, 0.0),
                  ('kl', 1, 'Closed', 30.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO', 'NO'])
    epc = loads([('l1', 1, 'Closed', 120.0, 0.0, 0.0, 0.0), ('ss', 1, 'Closed', 100.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO'])
    targets, reports = area.build_targets({'Load': epc, 'Gen': gens([])}, {'Load': case, 'Gen': gens([])})
    assert reports['area_load'].iloc[0]['Left As Is MW'] == 30
    est = area.loss_estimate(targets, reports['area_load'], pd.Series({1: 9.0}))
    assert abs(est['Est. Loss MW'].iloc[0] - 9.0 * (250.0 / 280.0) ** area.LOSS_EXPONENT) < 1e-9

def test_prebalance_moves_by_the_one_to_one_loads_change():
    # ss goes 50 -> 150 MW (EPC): +100 MW of load, so the gens go up 100. The conforming load and the gen keep their EPC values.
    case = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 0.0), ('ss', 1, 'Closed', 50.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO'], DistMWInput=0.0)
    epc = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 0.0), ('ss', 1, 'Closed', 150.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO'], DistMWInput=0.0)
    gen = gens([('g1', 1, '1', 'Closed', 'GT', 300.0, 1000.0, 0.0)])
    targets, _ = area.build_targets({'Load': epc, 'Gen': gen}, {'Load': case, 'Gen': gen})
    moved = area.prebalance_targets(targets)
    assert moved is not None and targets['Gen'].loc['g1', 'MWSetPoint_Target'] == 400.0
    assert targets['Load'].loc['ss', 'SMW_Target'] == 150.0 and targets['Load'].loc['l1', 'SMW_Target'] == 200.0

if __name__ == '__main__':
    tests = [f for name, f in dict(globals()).items() if name.startswith('test_')]
    for test in tests:
        test()
    print(f'{len(tests)} tests passed.')

def test_shortfall_sheet_lists_epc_and_pwb_gens_side_by_side():
    # Area 1: 1000 MW of EPC wind, no PWB wind (flagged first). Solar: 400 MW asked, the PWB's s1 + s2 take 150 (short
    # 250). s1 is in both cases: on one row. Area 2: Hydro 50 MW short, under the 100 MW threshold: not listed.
    case = gens([('s1', 1, '1', 'Closed', 'PV', 50.0, 100.0, 0.0), ('s2', 1, '2', 'Open', 'PV', 0.0, 50.0, 0.0),
                 ('t1', 1, '3', 'Closed', 'GT', 500.0, 3000.0, 0.0), ('h1', 2, '1', 'Closed', 'HY', 100.0, 100.0, 0.0)])
    epc = gens([('w1', 1, '7', 'Closed', 'WT', 600.0, 700.0, 0.0), ('w2', 1, '8', 'Closed', 'WT', 400.0, 500.0, 0.0),
                ('e9', 1, '9', 'Closed', 'PV', 300.0, 300.0, 0.0), ('s1', 1, '1', 'Closed', 'PV', 100.0, 100.0, 0.0),
                ('wo', 1, '6', 'Open', 'WT', 0.0, 900.0, 0.0),
                ('t1', 1, '3', 'Closed', 'GT', 500.0, 3000.0, 0.0), ('h1', 2, '1', 'Closed', 'HY', 150.0, 150.0, 0.0)])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case})
    s = reports['shortfall']
    assert s[['AreaNumber', 'GenType', 'Flag']].drop_duplicates().values.tolist() == [
        [1, 'Wind', 'no PWB units'], [1, 'Solar', 'short of capacity']]
    wind = s[s['GenType'] == 'Wind']
    assert wind['EPC ID'].tolist() == ['7', '8'] and wind['PWB ID'].isna().all()      # Closed EPC units only.
    assert (wind['Unmet MW'] == 1000).all() and (wind['PWB Units'] == 0).all()
    solar = s[s['GenType'] == 'Solar']
    assert solar['EPC ID'].tolist() == ['1', '9'] and solar['PWB ID'].tolist() == ['1', '2']    # s1 on the same row.
    assert solar['PWB Seed Status'].tolist() == ['Closed', 'Open'] and solar['Planned MW'].tolist() == [100.0, 50.0]
    assert (solar['Unmet MW'] == 250).all() and (solar['PWB Group MWMax'] == 150).all()
    with saved_settings():
        area.SHORTFALL_REPORT_MW = 2000.0
        _, reports = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case})
    assert reports['shortfall'].empty and 'PWB ID' in reports['shortfall']
