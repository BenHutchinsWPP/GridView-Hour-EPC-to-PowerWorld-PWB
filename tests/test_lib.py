"""
Tests for the parts of Scripts/ which don't need PowerWorld. Runs on any OS.
    python tests/test_lib.py      (or: pytest tests)
"""
import re
import sys
from pathlib import Path
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from Scripts import pw, prep, scale, area, settings, topo

class FakeSimAuto:
    """
    In-memory stand-in for SimAuto. Records written rows by (table, first field value).
    SolvePowerFlow fails while any row with a key in `bad` is applied.
    Reports two healthy buses (1 and 2) at 1.0 PU, and no shunts or branches.
    """
    def __init__(self, bad=()):
        self.bad = set(bad)
        self.applied = {}
        self.states = {}
        self.CreateIfNotFound = False

    def RunScriptCommand(self, command):
        name = re.search(r'"(.+?)"', command)
        if command.startswith('StoreState'):
            self.states[name.group(1)] = dict(self.applied)
        elif command.startswith('RestoreState'):
            self.applied = dict(self.states[name.group(1)])
        elif command.startswith('DeleteState'):
            del self.states[name.group(1)]
        elif command.startswith('SolvePowerFlow') and self.diverged():
            return ('Diverged',)
        return ('',)

    def diverged(self):
        return bool(self.bad & {k for _, k in self.applied})

    def ChangeParametersMultipleElementRect(self, table, fields, rows):
        for row in rows:
            self.applied[(table, row[0])] = dict(zip(fields, row))
        return ('',)

    def GetParametersMultipleElementRect(self, table, fields, filter_group):
        if table != 'Bus':
            return ('', [])
        values = {'Vpu': '1.0', 'IslandNumber': '1', 'BusIsStarBus:1': 'NO', 'BusNomVolt': '230'}
        return ('', [[str(n) if f in ('Busnum', 'Number') else values.get(f, '0') for f in fields] for n in (1, 2)])

def targets_df(**cols):
    df = pd.DataFrame(cols).set_index('ObjectID')
    df['ExclusionReason'] = ''
    return df

def scaling_targets(gen_status):
    """g1 on bus 1 and g2 on bus 2 scaling up; l1 scaling up and l2 closing in."""
    gen = targets_df(ObjectID=['g1', 'g2'], BusNum=[1, 2], ID='1', Status=gen_status, Status_Target='Closed',
                     MWSetPoint=[10.0, 10.0], MWSetPoint_Target=[50.0, 20.0])
    load = targets_df(ObjectID=['l1', 'l2'], BusNum=[1, 1], ID=['1', '2'], Status=['Closed', 'Open'], Status_Target='Closed',
                      DistStatus='Open', DistStatus_Target='Open',
                      **{c: [5.0, 0.0] for c in scale.SCALED['Load']}, **{f'{c}_Target': [10.0, 8.0] for c in scale.SCALED['Load']})
    pvqv = pd.DataFrame({'Number': [1, 2], 'Name': 'x', 'NomkV': 230.0, 'SensdVdPself': 0.0, 'SensdVdQself': 0.0})
    return {'Gen': gen, 'Load': load}, pvqv

# ------------------ pw ------------------
def test_set_df_bisect_returns_failures_and_rolls_them_back():
    sa = FakeSimAuto(bad={'b2', 'b5'})
    df = pd.DataFrame({'ObjectID': [f'b{i}' for i in range(8)], 'Status': 'Open'})
    failed = pw.set_df_bisect(sa, 'Branch', df, ['ObjectID', 'Status'])
    assert sorted(df.loc[failed, 'ObjectID']) == ['b2', 'b5']
    assert {k for _, k in sa.applied} == {'b0', 'b1', 'b3', 'b4', 'b6', 'b7'}
    assert sa.states == {}  # Every named state cleaned up.

def test_trial_rolls_back_on_exception():
    sa = FakeSimAuto()
    try:
        with pw.trial(sa):
            pw.set_df(sa, 'Gen', pd.DataFrame({'BusNum': [1]}), ['BusNum'])
            raise RuntimeError
    except RuntimeError:
        pass
    assert sa.applied == {}

def test_chk_raises():
    try:
        pw.chk(('Field not found',), 'Get Bus')
        assert False
    except pw.SimAutoError as e:
        assert 'Field not found' in str(e)

def test_set_df_writes_only_requested_fields():
    sa = FakeSimAuto()
    df = targets_df(ObjectID=['g1'], BusNum=[10], ID=['1'], Status=['Closed'], MWSetPoint=[5.0], Status_Target=['Open'])
    pw.set_df(sa, 'Gen', df, scale.write_fields('Gen'))
    assert sa.applied[('Gen', '10')] == {'BusNum': '10', 'ID': '1', 'Status': 'Closed', 'MWSetPoint': '5.0'}

# ------------------ prep ------------------
def test_suggested_shunt_status():
    df = pd.DataFrame({'MvarNom': [50, 50, -50, -50, 50], 'Status': ['Open', 'Closed', 'Closed', 'Open', 'Open'], 'Vpu': [0.9, 1.1, 0.9, 1.1, 1.0]})
    assert prep.suggested_shunt_status(df, 0.92, 1.08).tolist() == ['Closed', 'Open', 'Open', 'Closed', 'Open']

# ------------------ topo ------------------
def test_transformer_is_okay_flags_low_fixed_taps():
    df = pd.DataFrame({'Tap': [1.0, 1.0], 'NomkVFrom': 230.0, 'XFNomkVbaseFrom': 230.0, 'NomkVTo': 115.0, 'XFNomkVbaseTo': 115.0,
                       'TapFixedFrom': [1.0, 0.5], 'TapFixedTo': 1.0})
    assert topo.transformer_is_okay(df).tolist() == [True, False]

def test_branch_status_changes():
    def case(rows):
        df = pd.DataFrame(rows, columns=['ObjectID', 'Status', 'BranchDeviceType'])
        return {'Branch': {'df': df}, 'Transformer': {'df': df.iloc[0:0]}}
    right = case([['a', 'Closed', 'Line'], ['b', 'Closed', 'Line'], ['c', 'Closed', 'Line'], ['f', 'Closed', 'Fuse']])
    left = case([['a', 'Closed', 'Line'], ['b', 'Open', 'Line']])
    _, changes = topo.branch_status_changes(left, right)
    assert changes.set_index('ObjectID')['StatusLeft'].to_dict() == {'b': 'Open', 'c': 'Open'}  # Fuse ignored.

def test_build_targets():
    common = dict(BusName='A', NomkV=230.0)
    case_gen = pd.DataFrame({'ObjectID': ['g1', 'g2', 'g3'], 'BusNum': [1, 2, 3], 'ID': '1', 'Status': ['Closed', 'Open', 'Closed'], 'MWSetPoint': [10.0, 7.0, 5.0], **common})
    gv_gen = pd.DataFrame({'ObjectID': ['g1', 'g2'], 'BusNum': [1, 2], 'ID': '1', 'Status': ['Closed', 'Open'], 'MWSetPoint': [20.0, 9.0], **common})
    load = lambda ids: pd.DataFrame({'ObjectID': ids, 'BusNum': 1, 'ID': '1', 'Status': 'Closed', 'SMW': 50.0, 'SMvar': 5.0, 'IMW': 0.0, 'IMvar': 0.0, 'ZMW': 0.0, 'ZMvar': 0.0, 'DistStatus': 'Closed', 'DistMWInput': 3.0, 'DistMvarInput': 1.0, **common})
    targets = topo.build_targets({'Gen': gv_gen, 'Load': load(['l1'])}, {'Gen': case_gen, 'Load': load(['l1', 'l2'])})
    gen, ld = targets['Gen'], targets['Load']

    assert gen.loc['g1', 'MWSetPoint_Target'] == 20 and gen.loc['g1', 'MWSetPoint_Seed'] == 10
    assert gen.loc['g2', 'MWSetPoint'] == 0 and gen.loc['g2', 'MWSetPoint_Target'] == 0            # Open now and in target.
    assert gen.loc['g3', 'Status_Target'] == 'Open' and gen.loc['g3', 'MWSetPoint_Target'] == 0    # Missing from target.
    assert ld.loc['l1', 'DistMWInput_Target'] == 0 and ld.loc['l1', 'DistStatus_Target'] == 'Open'  # GridView dist gen unused.
    assert ld.loc['l1', 'DistMWInput'] == 3.0                                                       # Current dist gen kept.
    assert ld.loc['l2', 'Status_Target'] == 'Open' and ld.loc['l2', 'SMW_Target'] == 0
    assert scale.included(gen).all() and scale.included(ld).all()

# ------------------ scale ------------------
def test_exclude_keeps_first_reason():
    df = targets_df(ObjectID=['g1', 'g2'], BusNum=[1, 2])
    scale.exclude(df, df.index == 'g1', 'Swing Unit')
    scale.exclude(df, [True, True], 'PVQV')
    assert df['ExclusionReason'].tolist() == ['Swing Unit', 'PVQV']
    assert not scale.included(df).any()

def test_voltage_ok_and_pvqv():
    gen = targets_df(ObjectID=['g1', 'g2', 'g3'], BusNum=[1, 2, 3], MWSetPoint=[0.0, 0.0, 0.0], MWSetPoint_Target=[100.0, 1.0, 0.0])
    vpu = pd.Series({1: 1.0, 2: 0.8})
    assert scale.voltage_ok(gen, vpu).tolist() == [True, False, False]  # Bus 3 unknown -> excluded.

    load = targets_df(ObjectID=['l1'], BusNum=[1], SMW=[0.0], SMW_Target=[0.0], SMvar=[0.0], SMvar_Target=[0.0])
    pvqv = pd.DataFrame({'Number': [1, 2, 3], 'Name': 'x', 'NomkV': 230.0, 'SensdVdPself': 0.002, 'SensdVdQself': 0.0})
    result = scale.compute_pvqv_exclusions({'Gen': gen, 'Load': load}, pvqv)
    assert result['Gen']['Number'].tolist() == [1] and len(result['Load']) == 0

def test_mark_start_and_at_progress():
    gen = targets_df(ObjectID=['g1', 'g2'], MWSetPoint=[10.0, 0.0], MWSetPoint_Target=[50.0, 100.0])
    scale.mark_start({'Gen': gen})
    new_df = scale.at_progress('Gen', gen, pd.Series({'g1': 25, 'g2': 50}))
    assert new_df['MWSetPoint'].to_dict() == {'g1': 20.0, 'g2': 50.0}
    assert new_df['Progress'].to_dict() == {'g1': 25, 'g2': 50}
    assert gen['MWSetPoint'].tolist() == [10.0, 0.0]  # Unchanged until applied.

def test_close_all_related_gen_load_skips_excluded_and_restores_failures():
    sa = FakeSimAuto(bad={'2'})  # Closing the gen on bus 2 diverges.
    gen = targets_df(ObjectID=['g1', 'g2', 'g3'], BusNum=[1, 2, 3], ID='1', Status='Open', Status_Target='Closed', MWSetPoint=[0.0, 0.0, 0.0])
    gen.loc['g3', 'ExclusionReason'] = 'PVQV'
    load = targets_df(ObjectID=['l1'], BusNum=[9], ID=['1'], Status=['Open'], Status_Target=['Closed'], DistStatus=['Closed'], DistStatus_Target=['Closed'],
                      **{c: [1.0] for c in scale.SCALED['Load']})
    scale.close_all_related_gen_load(sa, {'Gen': gen, 'Load': load})
    assert gen['Status'].tolist() == ['Closed', 'Open', 'Open']
    assert gen['ExclusionReason'].tolist() == ['', 'Diverged closing at 0 MW', 'PVQV']
    assert ('Gen', '3') not in sa.applied
    assert load.loc['l1', 'Status'] == 'Closed' and load.loc['l1', 'SMW'] == 0

def results(log):
    return log['step_df']['Result'].tolist()

def test_iterate_grows_step_while_solving():
    sa = FakeSimAuto()
    targets, pvqv = scaling_targets(gen_status='Closed')
    log = scale.iterate_to_gen_load_targets(sa, targets, pvqv)
    assert log['step_df'][['From', 'To']].values.tolist() == [[0, 25], [25, 75], [75, 100]]
    assert results(log) == ['Solved'] * 3
    assert targets['Gen']['MWSetPoint'].tolist() == [50.0, 20.0]
    assert targets['Load']['SMW'].tolist() == [10.0, 8.0] and targets['Load'].loc['l2', 'Status'] == 'Closed'
    assert sa.states == {}

def test_iterate_excludes_gens_which_diverge_closing_in():
    sa = FakeSimAuto(bad={'2'})  # Closing the gen on bus 2 diverges.
    targets, _ = scaling_targets(gen_status=['Closed', 'Open'])
    log = scale.iterate_to_gen_load_targets(sa, targets)
    gen = targets['Gen']
    assert results(log) == ['Solved'] * 3
    assert gen.loc['g1', 'MWSetPoint'] == 50.0 and gen.loc['g2', 'Status'] == 'Open'
    assert log['gen']['ExclusionReason'].to_dict() == {'g2': 'Diverged closing at 0 MW'}

def test_iterate_halves_step_then_bisects_and_holds():
    sa = FakeSimAuto(bad={'2'})  # g2 is already closed, so it first diverges during scaling, and never recovers.
    targets, _ = scaling_targets(gen_status='Closed')
    log = scale.iterate_to_gen_load_targets(sa, targets)
    gen = targets['Gen']
    assert results(log)[:5] == ['Diverged, halved step'] * 4 + ['Bisected']  # 25, 12, 6, 3 halve; 1 is bisected.
    assert log['step_df']['To'].iloc[-1] == 100
    assert gen['MWSetPoint'].tolist() == [50.0, 10.0]  # g2 left where it got to.
    assert log['gen']['ExclusionReason'].to_dict() == {'g2': 'Diverged during scaling'}
    assert targets['Load']['SMW'].tolist() == [10.0, 8.0]
    assert sa.states == {}

class LaterConvergesSimAuto(FakeSimAuto):
    """Raising g2 (bus 2) above its start diverges, until g1 (bus 1) reaches 30 MW."""
    def diverged(self):
        mw = {k: float(v['MWSetPoint']) for (table, k), v in self.applied.items() if table == 'Gen'}
        return mw.get('2', 10.0) > 10.0 and mw.get('1', 10.0) < 30.0

def test_iterate_held_gen_resumes_once_it_converges():
    sa = LaterConvergesSimAuto()
    targets, _ = scaling_targets(gen_status='Closed')
    log = scale.iterate_to_gen_load_targets(sa, targets)
    gen = targets['Gen']
    assert 'Bisected' in results(log) and log['step_df']['CaughtUp'].sum() > 0
    assert gen['MWSetPoint'].tolist() == [50.0, 20.0]
    assert (gen['ExclusionReason'] == '').all() and (gen['HeldReason'] == '').all()
    assert sa.states == {}

# ------------------ area ------------------
def test_classify_gen_type():
    df = pd.DataFrame({
        'UnitTypeCode': ['HY', 'UN', 'PV (Photovoltaic)', 'OT', '', 'SC', 'ZZ', 'BA', 'w2', None],
        'FuelTypeCode': ['WAT', 'NG', '', 'SUN', '', 'NG', 'WND', 'MWH', '', None]})
    assert area.classify_gen_type(df).tolist() == [
        'Hydro', 'Thermal', 'Solar', 'Solar', 'Unknown', 'Unknown', 'Wind', 'Battery Storage', 'Wind', 'Unknown']

def loads(rows):
    """rows: (ObjectID, AreaNumber, Status, SMW, IMW, ZMW, SMvar)."""
    df = pd.DataFrame(rows, columns=['ObjectID', 'AreaNumber', 'Status', 'SMW', 'IMW', 'ZMW', 'SMvar'])
    return df.assign(BusNum=1, ID='1', IMvar=0.0, ZMvar=0.0, DistStatus='Closed', DistMWInput=5.0, DistMvarInput=1.0)

def test_load_targets():
    case = loads([('l1', 1, 'Closed', 60.0, 20.0, 20.0, 10.0), ('l2', 1, 'Closed', 100.0, 0.0, 0.0, 10.0),
                  ('lo', 1, 'Open', 50.0, 0.0, 0.0, 0.0),
                  ('m1', 2, 'Closed', 10.0, 0.0, 0.0, 1.0), ('m2', 2, 'Closed', 10.0, 0.0, 0.0, 1.0),
                  ('n1', 3, 'Closed', 10.0, 0.0, 0.0, 1.0)])
    epc = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 30.0), ('X1', 1, 'Closed', 100.0, 0.0, 0.0, 10.0),
                 ('m1', 2, 'Closed', 30.0, 0.0, 0.0, 3.0), ('m2', 2, 'Open', 99.0, 0.0, 0.0, 9.0)])
    df, report = area.load_targets(epc, case)
    assert 'lo' not in df.index                                              # Open loads aren't scaled.
    assert df['SMW_Target'].to_dict() == {'l1': 150.0, 'l2': 150.0, 'm1': 30.0, 'm2': 0.0, 'n1': 0.0}
    assert df.loc[['l1', 'l2'], 'SMvar_Target'].tolist() == [20.0, 20.0]
    assert (df[[f'{c}_Target' for c in ['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']]] == 0).all().all()
    assert (df['DistStatus_Target'] == 'Open').all() and (df['Status_Target'] == df['Status']).all()
    assert df['Method'].to_dict() == {'l1': 'Area scale', 'l2': 'Area scale', 'm1': 'Matched keys', 'm2': 'Matched keys', 'n1': 'Area scale'}
    r = report.set_index('AreaNumber')
    assert r.loc[1, 'Target MW'] == 300 and r.loc[1, 'Seed MW'] == 200 and r.loc[1, 'EPC-only MW'] == 100
    assert r.loc[2, 'Method'] == 'Matched keys' and r.loc[3, 'Target MW'] == 0
    assert (r['Unmet MW'].abs() < 1e-9).all()

def test_load_targets_area_with_no_seed_load_is_unmet():
    case = loads([('l1', 1, 'Open', 10.0, 0.0, 0.0, 0.0)])
    epc = loads([('l1', 1, 'Closed', 40.0, 0.0, 0.0, 0.0)])
    df, report = area.load_targets(epc, case)
    assert len(df) == 0 and report.set_index('AreaNumber').loc[1, 'Unmet MW'] == 40

def gens(rows):
    """rows: (ObjectID, AreaNumber, ID, Status, UnitTypeCode, MWSetPoint, MWMax, MWMin)."""
    df = pd.DataFrame(rows, columns=['ObjectID', 'AreaNumber', 'ID', 'Status', 'UnitTypeCode', 'MWSetPoint', 'MWMax', 'MWMin'])
    return df.assign(BusNum=1, FuelTypeCode='')

def test_gen_targets():
    case = gens([('h1', 1, '1', 'Closed', 'HY', 10.0, 100.0, 0.0), ('h2', 1, '2', 'Closed', 'HY', 10.0, 300.0, 0.0),
                 ('w1', 1, '3', 'Closed', 'WT', 0.0, 50.0, 0.0), ('w2', 1, '4', 'Closed', 'WT', 0.0, 50.0, 0.0),
                 ('b1', 1, '5', 'Closed', 'BA', 0.0, 20.0, -20.0), ('b2', 1, '6', 'Closed', 'BA', 0.0, 60.0, -60.0),
                 ('t1', 1, '7', 'Closed', 'GT', 80.0, 100.0, 20.0), ('to', 1, '8', 'Open', 'GT', 0.0, 100.0, 0.0),
                 ('xx', 1, 'xx', 'Closed', 'BA', 500.0, 20000.0, -20000.0),
                 ('a', 2, '1', 'Closed', 'GT', 10.0, 100.0, 0.0), ('b', 2, '2', 'Closed', 'GT', 10.0, 100.0, 0.0)])
    epc = gens([('e1', 1, '1', 'Closed', 'HY', 200.0, 500.0, 0.0), ('e2', 1, '2', 'Closed', 'W3', 150.0, 200.0, 0.0),
                ('e3', 1, '3', 'Closed', 'ES', -40.0, 80.0, -80.0), ('e4', 1, '4', 'Closed', 'PV', 25.0, 30.0, 0.0),
                ('e5', 1, '5', 'Open', 'GT', 90.0, 100.0, 0.0),
                ('a', 2, '1', 'Closed', 'GT', 70.0, 100.0, 0.0), ('b', 2, '2', 'Open', 'GT', 50.0, 100.0, 0.0)])
    df, report = area.gen_targets(epc, case)
    assert 'xx' not in df.index and 'to' not in df.index                    # Swing, and gens offline before and after.
    # Hydro, wind, storage: all on at the same %. Thermal 0 MW: off. Matched area 2: EPC statuses and values.
    assert df['MWSetPoint_Target'].to_dict() == {'h1': 50.0, 'h2': 150.0, 'w1': 50.0, 'w2': 50.0, 'b1': -10.0, 'b2': -30.0,
                                                 't1': 0.0, 'a': 70.0, 'b': 0.0}
    assert df.loc[df['Status_Target'] == 'Open'].index.tolist() == ['t1', 'b']
    r = report.set_index(['AreaNumber', 'GenType'])['Unmet MW']
    assert r[(1, 'Wind')] == 50 and r[(1, 'Solar')] == 25 and r[(1, 'Hydro')] == 0 and r[(1, 'Thermal')] == 0

def test_load_targets_scale_positive_and_negative_separately():
    # Mvar +100 / -98 (net +2) against +300 / -50: one net factor would blow up; split by sign it doesn't.
    case = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 100.0), ('l2', 1, 'Closed', 50.0, 0.0, 0.0, -98.0),
                  ('dc', 1, 'Closed', -500.0, 0.0, 0.0, 0.0)])
    epc = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 300.0), ('e2', 1, 'Closed', -400.0, 0.0, 0.0, -50.0)])
    df, report = area.load_targets(epc, case)
    assert df['SMW_Target'].to_dict() == {'l1': 200.0, 'l2': 100.0, 'dc': -400.0}
    assert df['SMvar_Target'].to_dict() == {'l1': 300.0, 'l2': -50.0, 'dc': 0.0}
    r = report.iloc[0]
    assert (r['Target MW +'], r['Target MW -'], r['Seed MW +'], r['Seed MW -']) == (300, -400, 150, -500)
    assert all(abs(r[c]) < 1e-9 for c in report.columns if c.startswith('Unmet'))

def test_load_targets_mvar_spread_by_mw_when_no_positive_mvar():
    case = loads([('l1', 1, 'Closed', 10.0, 0.0, 0.0, 0.0), ('l2', 1, 'Closed', 20.0, 0.0, 0.0, 0.0)])
    epc = loads([('e1', 1, 'Closed', 30.0, 0.0, 0.0, 30.0)])
    df, _ = area.load_targets(epc, case)
    assert df['SMvar_Target'].tolist() == [10.0, 20.0]

def test_load_targets_negative_target_with_no_negative_seed_load_is_unmet():
    case = loads([('l1', 1, 'Closed', 10.0, 0.0, 0.0, 0.0)])
    epc = loads([('e1', 1, 'Closed', 20.0, 0.0, 0.0, 0.0), ('e2', 1, 'Closed', -30.0, 0.0, 0.0, 0.0)])
    df, report = area.load_targets(epc, case)
    assert df['SMW_Target'].tolist() == [20.0]
    assert report['Unmet MW -'].tolist() == [-30.0] and report['Unmet MW +'].tolist() == [0.0]

def test_gen_targets_leave_dc_ties_alone():
    case = gens([('d1', 1, '1', 'Closed', 'DC', 300.0, 500.0, -500.0), ('g1', 1, '2', 'Closed', 'GT', 0.0, 100.0, 0.0)])
    epc = gens([('d1', 1, '1', 'Closed', 'DC', -400.0, 500.0, -500.0), ('g2', 1, '2', 'Closed', 'GT', 60.0, 100.0, 0.0)])
    df, report = area.gen_targets(epc, case)
    assert df.index.tolist() == ['g1'] and df.loc['g1', 'MWSetPoint_Target'] == 60
    assert report['GenType'].tolist() == ['Thermal']

def test_achieved_reads_the_final_case_and_open_load_cleanup():
    case = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 10.0), ('lo', 1, 'Open', 50.0, 0.0, 0.0, 0.0)])
    epc = loads([('l1', 1, 'Closed', 80.0, 0.0, 0.0, 8.0)])
    case_gen = gens([('h1', 1, '1', 'Closed', 'HY', 10.0, 100.0, 0.0), ('xx', 1, 'xx', 'Closed', '', 900.0, 2e4, -2e4)]).assign(FuelTypeCode=['', 'MWH'])
    epc_gen = gens([('h1', 1, '1', 'Closed', 'HY', 40.0, 100.0, 0.0)])
    _, reports = area.build_targets({'Load': epc, 'Gen': epc_gen}, {'Load': case, 'Gen': case_gen})
    # The final case stopped short: 70 MW of load (the open load doesn't count), 30 MW of hydro, swing not counted.
    final = {'Load': case.assign(SMW=[70.0, 999.0]), 'Gen': case_gen.assign(MWSetPoint=[30.0, 900.0])}
    a = area.achieved(final, reports)
    load = a['area_load'].set_index('AreaNumber')
    assert load.loc[1, 'Final MW'] == 70 and load.loc[1, 'Final Error MW'] == -10
    assert load.loc[1, 'Final MW +'] == 70 and load.loc[1, 'Final Error MW +'] == -10 and load.loc[1, 'Final MW -'] == 0
    assert a['area_gen'][['GenType', 'Final MW', 'Final Error MW']].values.tolist() == [['Hydro', 30.0, -10.0]]
    cleanup = area.open_load_cleanup(case)
    assert cleanup['ObjectID'].tolist() == ['lo'] and cleanup['DistMWInput'].tolist() == [0.0]

def test_area_targets_through_the_engine():
    case = loads([('l1', 1, 'Closed', 60.0, 20.0, 20.0, 10.0), ('l2', 1, 'Closed', 100.0, 0.0, 0.0, 10.0)])
    epc = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 40.0)])
    case_gen = gens([('h1', 1, '1', 'Closed', 'HY', 10.0, 100.0, 0.0), ('h2', 1, '2', 'Closed', 'HY', 10.0, 300.0, 0.0)])
    epc_gen = gens([('e1', 1, '1', 'Closed', 'HY', 200.0, 500.0, 0.0)])
    targets, _ = area.build_targets({'Load': epc, 'Gen': epc_gen}, {'Load': case, 'Gen': case_gen})
    sa = FakeSimAuto()
    log = scale.iterate_to_gen_load_targets(sa, targets)
    assert results(log) == ['Solved'] * 3 and sa.states == {}
    load, gen = targets['Load'], targets['Gen']
    assert load['SMW'].tolist() == [150.0, 150.0] and load['SMvar'].tolist() == [20.0, 20.0]
    assert (load[['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']] == 0).all().all()
    assert (load['DistStatus'] == 'Open').all() and (load['Status'] == 'Closed').all()
    assert gen['MWSetPoint'].tolist() == [50.0, 150.0] and (gen['Status'] == 'Closed').all()
    assert sa.applied[('Load', '1')]['DistStatus'] == 'Open'

def test_swing_mw():
    gen = gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 2, '1', 'Closed', 'GT', 0.0, 100.0, 0.0)]).assign(MW=[120.0, 50.0], BusNum=[7, 8])
    assert area.swing_mw(gen, slack_buses=[8]) == (120.0, [7])                  # The giant swing wins.
    assert area.swing_mw(gen[gen['ID'] != 'xx'], slack_buses=[8]) == (50.0, [8])  # Else gens on slack buses.

def test_balance_targets_half_gen_half_load():
    # Swing at +100 MW: gens up 50 MW, load down 50 MW, each by one factor over the whole case.
    gen = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 300.0, 0.0), ('g2', 2, '2', 'Closed', 'HY', 100.0, 120.0, 0.0),
                ('neg', 2, '3', 'Closed', 'BA', -10.0, 50.0, -50.0), ('sw', 3, '1', 'Closed', 'GT', 500.0, 1e4, 0.0),
                ('dc', 3, '2', 'Closed', 'DC', 90.0, 500.0, -500.0)]).assign(BusNum=[1, 2, 2, 9, 3])
    load = loads([('l1', 1, 'Closed', 150.0, 0.0, 0.0, 30.0), ('l2', 2, 'Closed', 350.0, 0.0, 0.0, 70.0),
                  ('lneg', 2, 'Closed', -40.0, 0.0, 0.0, 0.0), ('lo', 2, 'Open', 99.0, 0.0, 0.0, 0.0)])
    t = area.balance_targets({'Gen': gen, 'Load': load}, swing=100.0, swing_buses=[9])
    assert t['Gen']['MWSetPoint_Target'].to_dict() == {'g1': 125.0, 'g2': 120.0}  # g2 capped at MWMax.
    assert t['Load']['SMW_Target'].to_dict() == {'l1': 135.0, 'l2': 315.0}
    assert t['Load']['SMvar_Target'].to_dict() == {'l1': 27.0, 'l2': 63.0}
    assert (t['Load']['Status_Target'] == 'Closed').all()

def test_thermal_scaling_up_brings_biggest_units_on_first():
    # Thermal target 250 MW with 100 MW online: g3 (online in the EPC) comes on before the bigger g2, and is enough.
    case = gens([('g1', 1, '1', 'Closed', 'GT', 50.0, 100.0, 0.0), ('g2', 1, '2', 'Open', 'GT', 0.0, 400.0, 0.0),
                 ('g3', 1, '3', 'Open', 'GT', 0.0, 200.0, 0.0), ('g4', 1, '4', 'Open', 'GT', 0.0, 50.0, 0.0),
                 ('sc', 1, '5', 'Closed', 'SC', 0.0, 0.0, 0.0)])
    epc = gens([('g3', 1, '3', 'Closed', 'GT', 150.0, 200.0, 0.0), ('e1', 1, '9', 'Closed', 'CC', 100.0, 300.0, 0.0)])
    df, report = area.gen_targets(epc, case)
    assert df['Status_Target'].to_dict() == {'g1': 'Closed', 'g3': 'Closed', 'sc': 'Closed'}
    assert df['MWSetPoint_Target'].round(6).to_dict() == {'g1': 83.333333, 'g3': 166.666667, 'sc': 0.0}
    r = report.set_index(['AreaNumber', 'GenType'])
    assert r.loc[(1, 'Thermal'), ['Committed', 'Decommitted', 'Unmet MW']].round(6).tolist() == [1, 0, 0]
    # Same, but nothing online in the EPC: the biggest (g2) comes on.
    df, _ = area.gen_targets(epc.iloc[[1]], case)
    assert df['Status_Target'].to_dict() == {'g1': 'Closed', 'g2': 'Closed', 'sc': 'Closed'}

def test_thermal_scaling_down_takes_smallest_units_off_first():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 100.0, 0.0), ('g2', 1, '2', 'Closed', 'ST', 200.0, 200.0, 0.0),
                 ('g3', 1, '3', 'Closed', 'CC', 300.0, 300.0, 0.0)])
    epc = gens([('e1', 1, '9', 'Closed', 'GT', 200.0, 300.0, 0.0)])
    df, _ = area.gen_targets(epc, case)
    assert df['Status_Target'].to_dict() == {'g1': 'Open', 'g2': 'Open', 'g3': 'Closed'}
    assert df['MWSetPoint_Target'].to_dict() == {'g1': 0.0, 'g2': 0.0, 'g3': 200.0}
    # g1 is online in the EPC: g2 (offline in the EPC) goes off first instead, and g1 stays.
    epc = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 100.0, 0.0), ('e1', 1, '9', 'Closed', 'GT', 100.0, 300.0, 0.0)])
    df, _ = area.gen_targets(epc, case)
    assert df['Status_Target'].to_dict() == {'g1': 'Closed', 'g2': 'Open', 'g3': 'Closed'}
    assert df['MWSetPoint_Target'].to_dict() == {'g1': 50.0, 'g2': 0.0, 'g3': 150.0}

def test_commitment_off_keeps_seed_statuses():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 50.0, 100.0, 0.0), ('g2', 1, '2', 'Open', 'GT', 0.0, 400.0, 0.0)])
    epc = gens([('e1', 1, '9', 'Closed', 'GT', 150.0, 300.0, 0.0)])
    area.COMMIT = False
    try:
        df, report = area.gen_targets(epc, case)
    finally:
        area.COMMIT = True
    assert df['Status_Target'].to_dict() == {'g1': 'Closed'} and df['MWSetPoint_Target'].tolist() == [100.0]
    assert report['Unmet MW'].tolist() == [50.0]

def test_commitment_headroom_and_all_on_types():
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
    assert st['h1'] == 'Closed' and df.loc[['h1', 'h2'], 'MWSetPoint_Target'].tolist() == [10.0, 30.0]  # Run of the river.
    assert st['s1'] == 'Open' and 's2' not in st           # Solar target 0: all off.
    assert df.loc['s1', 'MWSetPoint_Target'] == 0
    # Storage target 0: seed statuses, at 0 MW.
    case = gens([('b1', 1, '1', 'Closed', 'BA', 5.0, 50.0, -50.0), ('b2', 1, '2', 'Open', 'BA', 0.0, 50.0, -50.0)])
    df, _ = area.gen_targets(gens([]), case)
    assert df['Status_Target'].to_dict() == {'b1': 'Closed'} and df['MWSetPoint_Target'].tolist() == [0.0]

def test_build_targets_leaves_slack_bus_gens_out_of_both_cases():
    # Identical cases: the slack-bus unit's 300 MW is not loaded onto g1, and the area still matches by key.
    case = gens([('s1', 1, '1', 'Closed', 'GT', 300.0, 500.0, 0.0), ('g1', 1, '2', 'Closed', 'GT', 80.0, 100.0, 0.0)]).assign(BusNum=[5, 6])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': case}, {'Load': loads([]), 'Gen': case}, swing_buses=[5])
    assert targets['Gen'].index.tolist() == ['g1'] and targets['Gen'].loc['g1', 'MWSetPoint_Target'] == 80
    assert reports['area_gen']['Method'].tolist() == ['Matched keys'] and reports['area_gen']['Unmet MW'].tolist() == [0.0]

# ------------------ settings ------------------
def test_repo_settings_file_loads():
    saved = {(m, k): v for m in settings.MODULES.values() for k, v in vars(m).items() if k.isupper()}
    try:
        run = settings.load()
        assert run['mode'] in settings.MODES
    finally:
        for (m, k), v in saved.items():
            setattr(m, k, v)

def test_settings_apply_and_reject(tmp_path=None):
    import tempfile
    saved = (scale.EXCLUDE_VMIN, dict(area.UNIT_GENTYPE), set(area.ALL_ON_TYPES), list(area.GEN_TYPES))
    with tempfile.TemporaryDirectory() as d:
        fp = Path(d) / 's.toml'
        try:
            fp.write_text('[run]\nmode = "topology"\n[scale]\nexclude_vmin = 0.7\n'
                          '[area]\nall_on_types = ["Wind"]\ngen_types = ["Solar", "Hydro", "Wind", "Thermal", "Battery Storage", "Unknown", "Pumped"]\n'
                          '[area.unit_gentype]\nps = "Pumped"\nHY = "fuel"\n')
            run = settings.load(fp)
            assert run['mode'] == 'topology' and run['balance_swing'] is True     # Default filled in.
            assert scale.EXCLUDE_VMIN == 0.7 and area.ALL_ON_TYPES == {'Wind'}
            assert area.UNIT_GENTYPE['PS'] == 'Pumped' and area.UNIT_GENTYPE['HY'] is None and area.UNIT_GENTYPE['GT'] == 'Thermal'
            # The setting reaches code which reads it at call time.
            gen = targets_df(ObjectID=['g1'], BusNum=[1])
            assert scale.voltage_ok(gen, pd.Series({1: 0.75})).tolist() == [True]
            scale.EXCLUDE_VMIN, area.UNIT_GENTYPE, area.ALL_ON_TYPES, area.GEN_TYPES = saved[0], dict(saved[1]), set(saved[2]), list(saved[3])
            for bad in ['[run]\nmode = "x"\n', '[run]\nmodee = "area"\n', '[scale]\nnope = 1\n', '[nope]\n',
                        '[area.unit_gentype]\nPS = "Pumped"\n', '[run]\nswing_bus = true\n', '[scale]\nexclude_vmin = "low"\n',
                        '[run]\nbalance_swing = 1\n', '[scale]\nstatus = {}\n', '[area]\nall_on_types = "Wind"\n']:
                fp.write_text(bad)
                try:
                    settings.load(fp)
                    assert False, bad
                except ValueError:
                    pass
        finally:
            scale.EXCLUDE_VMIN, area.UNIT_GENTYPE, area.ALL_ON_TYPES, area.GEN_TYPES = saved

if __name__ == '__main__':
    tests = [f for name, f in dict(globals()).items() if name.startswith('test_')]
    for test in tests:
        test()
    print(f'{len(tests)} tests passed.')
