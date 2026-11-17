"""
Tests for the parts of Scripts/ which don't need PowerWorld. Runs on any OS.
    python tests/test_lib.py      (or: pytest tests)
"""
import contextlib
import importlib.util
import io
import sys
import tempfile
from contextlib import contextmanager
from pathlib import Path
import numpy as np
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
        self.checkpoint = None
        self.CreateIfNotFound = False

    def SaveState(self):
        self.checkpoint = dict(self.applied)
        return ('',)

    def LoadState(self):
        self.applied = dict(self.checkpoint)
        return ('',)

    def RunScriptCommand(self, command):
        if command.startswith('SolvePowerFlow') and self.diverged():
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

LOAD_P_ALL = scale.SCALED['Load']

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
def test_statcom_skips_buses_which_already_have_one():
    class LowVoltageSimAuto(FakeSimAuto):
        """Buses 1 and 2 at 0.5 PU (bus 1 lowest); bus 1 already has a STATCOM from an earlier run."""
        def GetParametersMultipleElementRect(self, table, fields, filter_group):
            if table == 'Gen':
                return ('', [[{'BusNum': '1', 'ID': scale.STATCOM_ID}[f] for f in fields]])
            if table != 'Bus' or 'Vpu' not in fields:
                return super().GetParametersMultipleElementRect(table, fields, filter_group)
            values = {'Number': ['1', '2'], 'BusNomVolt': ['230', '230'], 'Vpu': ['0.4', '0.5'], 'IslandNumber': ['1', '1'], 'BusIsStarBus:1': ['NO', 'NO']}
            return ('', [[values[f][i] for f in fields] for i in range(2)])
    sa = LowVoltageSimAuto()
    assert scale.create_statcom_on_lowestv_bus(sa, skip_buses=[]) == (2, True)
    assert sa.applied[('Gen', '2')]['GenStatus'] == 'Closed'                    # Created Closed, in one trial.

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

class BigStepDivergesSimAuto(FakeSimAuto):
    """Moving g1 (bus 1) by more than 15 MW in one write diverges; g1 goes 10 -> 50 MW, so steps above 37% fail."""
    big = False
    def ChangeParametersMultipleElementRect(self, table, fields, rows):
        for row in rows:
            if table == 'Gen' and row[0] == '1' and 'MWSetPoint' in fields:
                prev = float(self.applied.get(('Gen', '1'), {}).get('MWSetPoint', 10.0))
                self.big |= abs(float(row[fields.index('MWSetPoint')]) - prev) > 15.0
        return super().ChangeParametersMultipleElementRect(table, fields, rows)
    def diverged(self):
        return self.big or super().diverged()
    def LoadState(self):
        self.big = False
        return super().LoadState()

def test_iterate_keeps_the_step_after_a_halved_success():
    # 25 solves; doubled to 50 diverges; 25 solves once halved, so the next step stays 25 (not 50, which would fail).
    log = scale.iterate_to_gen_load_targets(BigStepDivergesSimAuto(), scaling_targets(gen_status='Closed')[0])
    assert log['step_df'][['From', 'To']].values.tolist() == [[0, 25], [25, 75], [25, 50], [50, 75], [75, 100]]
    assert results(log) == ['Solved', 'Diverged, halved step', 'Solved', 'Solved', 'Solved']

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
    return df.assign(BusNum=1, ID='1', AGC='YES', IMvar=0.0, ZMvar=0.0, DistStatus='Closed', DistMWInput=5.0, DistMvarInput=1.0)

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
    # Hydro, wind, storage: all on at the same %. Thermal 0 MW: off. Matched area 2: EPC statuses and values.
    assert df['MWSetPoint_Target'].to_dict() == {'h1': 50.0, 'h2': 150.0, 'w1': 50.0, 'w2': 50.0, 'b1': -10.0, 'b2': -30.0,
                                                 't1': 0.0, 'a': 70.0, 'b': 0.0}
    assert df.loc[df['Status_Target'] == 'Open'].index.tolist() == ['t1', 'b']
    r = report.set_index(['AreaNumber', 'GenType'])['Unmet MW']
    assert r[(1, 'Wind')] == 50 and r[(1, 'Solar')] == 25 and r[(1, 'Hydro')] == 0 and r[(1, 'Thermal')] == 0
    # With make-up: Wind's 50 MW (capped) and Solar's 25 MW (no PWB units) go onto area 1's Hydro, by MW (1:3).
    # Storage (charging) takes none; area 2 is untouched. Unmet still says what each bucket's own units couldn't take.
    df, report = area.gen_targets(epc, case)
    assert df['MWSetPoint_Target'].to_dict() == {'h1': 68.75, 'h2': 206.25, 'w1': 50.0, 'w2': 50.0, 'b1': -10.0, 'b2': -30.0,
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

def test_area_targets_through_the_engine():
    case = loads([('l1', 1, 'Closed', 60.0, 20.0, 20.0, 10.0), ('l2', 1, 'Closed', 100.0, 0.0, 0.0, 10.0)])
    epc = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 40.0)])
    case_gen = gens([('h1', 1, '1', 'Closed', 'HY', 0.0, 100.0, 0.0), ('h2', 1, '2', 'Closed', 'HY', 0.0, 300.0, 0.0)])
    epc_gen = gens([('e1', 1, '1', 'Closed', 'HY', 200.0, 500.0, 0.0)])
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

def test_swing_mw():
    gen = gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 2, '1', 'Closed', 'GT', 0.0, 100.0, 0.0)]).assign(MW=[120.0, 50.0], BusNum=[7, 8])
    assert area.swing_mw(gen, slack_buses=[8]) == (120.0, [7])                  # The giant swing wins.
    assert area.swing_mw(gen[gen['ID'] != 'xx'], slack_buses=[8]) == (50.0, [8])  # Else gens on slack buses.

def test_balance_targets_moves_gen_only():
    # Swing at +100 MW: gens up 100 MW by one factor (50 each); load untouched. g2 (area 2) caps at MWMax 120 and area 2
    # has no other gen, so its last 30 MW go onto g1.
    gen = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 300.0, 0.0), ('g2', 2, '2', 'Closed', 'HY', 100.0, 120.0, 0.0),
                ('neg', 2, '3', 'Closed', 'BA', -10.0, 50.0, -50.0), ('sw', 3, '1', 'Closed', 'GT', 500.0, 1e4, 0.0),
                ('dc', 3, '2', 'Closed', 'DC', 90.0, 500.0, -500.0)]).assign(BusNum=[1, 2, 2, 9, 3])
    load = loads([('l1', 1, 'Closed', 150.0, 0.0, 0.0, 30.0), ('l2', 2, 'Closed', 350.0, 0.0, 0.0, 70.0),
                  ('lneg', 2, 'Closed', -40.0, 0.0, 0.0, 0.0), ('lo', 2, 'Open', 99.0, 0.0, 0.0, 0.0)])
    t = area.balance_targets({'Gen': gen, 'Load': load}, swing=100.0, swing_buses=[9])
    assert t['Gen']['MWSetPoint_Target'].to_dict() == {'g1': 180.0, 'g2': 120.0}
    assert t['Load'].empty and (t['Gen']['MWSetPoint_Target'] - t['Gen']['MWSetPoint_Seed']).sum() == 100.0
    area.rescale_targets({'Gen': targets_df(ObjectID=['g1'], MWSetPoint=[150.0], MWSetPoint_Target=[150.0]),
                          'Load': targets_df(ObjectID=['l1'], **{c: [1.0] for c in area.LOAD_P + area.LOAD_Q})}, t)   # Empty load table is fine.

def test_spread_by_group_fills_the_group_then_its_area_then_everyone():
    # Area 1: Thermal t1 (100, cap 110), Hydro h1 (100, cap 1000). Area 2: Thermal t2 (100, cap 1000).
    mw = pd.Series([100.0, 100.0, 100.0], index=['t1', 'h1', 't2'])
    cap = pd.Series([110.0, 1000.0, 1000.0], index=mw.index)
    area_, group = pd.Series([1, 1, 2], index=mw.index), pd.Series(['Thermal', 'Hydro', 'Thermal'], index=mw.index)
    ask = pd.Series({(1, 'Thermal'): 50.0})
    x, left = area.spread_by_group(ask, mw, cap, area_, group, spill_all=False)
    assert x.to_dict() == {'t1': 10.0, 'h1': 40.0, 't2': 0.0} and left == 0.0       # Group first, then its area.
    cap['h1'] = 120.0
    x, left = area.spread_by_group(ask, mw, cap, area_, group, spill_all=False)
    assert x.to_dict() == {'t1': 10.0, 'h1': 20.0, 't2': 0.0} and left == 20.0      # The area is full: stays short.
    x, left = area.spread_by_group(ask, mw, cap, area_, group)
    assert x.to_dict() == {'t1': 10.0, 'h1': 20.0, 't2': 20.0} and left == 0.0      # spill_all: everyone else.
    x, left = area.spread_by_group(pd.Series({(1, 'Thermal'): -150.0}), mw, cap, area_, group, spill_all=False)
    assert x.to_dict() == {'t1': -100.0, 'h1': -50.0, 't2': 0.0} and left == 0.0   # Down: floor 0, then the area.
    x, _ = area.spread_evenly(-30.0, mw, cap, area_, group)
    assert x.round(9).to_dict() == {'t1': -10.0, 'h1': -10.0, 't2': -10.0}        # One factor on every group.
    x, left = area.spread_evenly(50.0, mw.iloc[:0], cap.iloc[:0], area_.iloc[:0], group.iloc[:0])
    assert left == 50.0                                                             # Nothing to place it on.

def test_spread_by_group_goes_by_mw_and_keeps_opposite_moves_apart():
    # By MW, not per group: t (100) and h (300) share +40 as 10 / 30.
    mw = pd.Series([100.0, 300.0], index=['t', 'h'])
    x, _ = area.spread_evenly(40.0, mw, pd.Series(1e4, index=mw.index), pd.Series(1, index=mw.index), pd.Series(['Thermal', 'Hydro'], index=mw.index))
    assert x.round(9).to_dict() == {'t': 10.0, 'h': 30.0}
    # Thermal (at MWMax) asks +50, Hydro asks -30, Wind has room. Thermal's 50 go onto Wind only, never back onto Hydro.
    mw = pd.Series([100.0, 100.0, 100.0], index=['t1', 'h1', 'w1'])
    cap = pd.Series([100.0, 1000.0, 1000.0], index=mw.index)
    group = pd.Series(['Thermal', 'Hydro', 'Wind'], index=mw.index)
    x, left = area.spread_by_group(pd.Series({(1, 'Thermal'): 50.0, (1, 'Hydro'): -30.0}), mw, cap,
                                   pd.Series(1, index=mw.index), group, spill_all=False)
    assert x.round(9).to_dict() == {'t1': 0.0, 'h1': -30.0, 'w1': 50.0} and left == 0.0

def test_balance_skips_held_and_out_of_band_units():
    # FakeSimAuto has buses 1 and 2 at 1.0 PU; bus 3 has no voltage, so it counts as outside the band.
    gen = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 300.0, 0.0), ('g2', 1, '2', 'Closed', 'GT', 100.0, 300.0, 0.0),
                ('g3', 1, '3', 'Closed', 'GT', 100.0, 300.0, 0.0)]).assign(BusNum=[1, 2, 3])
    load = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 0.0), ('l2', 1, 'Closed', 100.0, 0.0, 0.0, 0.0)])
    main = {'Gen': gen.set_index('ObjectID').assign(ExclusionReason='', HeldReason=['', 'Voltage', '']),
            'Load': load.set_index('ObjectID').assign(ExclusionReason=['', 'PVQV'], HeldReason='')}
    skip = area.balance_skips(FakeSimAuto(), {'Gen': gen, 'Load': load}, main)
    assert skip == {'g2', 'g3', 'l2'}
    t = area.balance_targets({'Gen': gen, 'Load': load}, swing=100.0, swing_buses=[], skip=skip)
    assert t['Gen']['MWSetPoint_Target'].to_dict() == {'g1': 200.0}     # The whole swing lands on what is left.
    assert t['Load'].empty

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

class SwingSimAuto(FakeSimAuto):
    """
    A case with gens & loads on their own buses, and the giant swing (ID xx, bus 99) carrying load minus generation.
    `slack`: bus numbers which report Slack = YES. Gens on them carry the imbalance too (so they are the swing when there is no xx).
    Written rows (keyed by bus) overlay the starting tables; StoreState/RestoreState work on them as in FakeSimAuto.
    """
    def __init__(self, gen, load, slack=()):
        super().__init__()
        self.slack = set(slack)
        self.base = {'Gen': gen.set_index('BusNum', drop=False), 'Load': load.set_index('BusNum', drop=False)}

    def table(self, name):
        df = self.base[name].copy()
        for (t, key), row in self.applied.items():
            if t == name:
                for f, v in row.items():
                    if f in df.columns:
                        df.loc[int(key), f] = type(df[f].iloc[0])(v) if f not in ('Status', 'DistStatus', 'ID') else v
        if name == 'Gen':
            on = df['Status'] == 'Closed'
            ld = self.table('Load')
            p = ld.loc[ld['Status'] == 'Closed', ['SMW', 'IMW', 'ZMW']].sum().sum()
            if getattr(self, 'count_dist', False):     # Opt-in: dist gen injects while its load and itself are Closed.
                p -= ld.loc[(ld['Status'] == 'Closed') & (ld['DistStatus'] == 'Closed'), 'DistMWInput'].sum()
            df['MW'] = df['MWSetPoint']
            sw = (df['ID'] == 'xx') | df['BusNum'].isin(self.slack)
            df.loc[sw, 'MW'] = p - df.loc[on & ~sw, 'MWSetPoint'].sum()
        return df

    def GetParametersMultipleElementRect(self, table, fields, filter_group):
        if table in self.base:
            return ('', [[str(v) for v in row] for row in self.table(table)[fields].itertuples(index=False)])
        if table == 'Area':
            return ('', [[str(n), str(mw)] for n, mw in getattr(self, 'losses', {}).items()])
        if table == 'Bus':
            buses = sorted(set(self.base['Gen']['BusNum']) | set(self.base['Load']['BusNum']))
            values = {'Vpu': '1.0', 'IslandNumber': '1', 'BusIsStarBus:1': 'NO', 'BusNomVolt': '230'}
            islands = getattr(self, 'islands', {})      # Bus -> IslandNumber text ('' = blank); 1 if not given.
            return ('', [[str(n) if f in ('Busnum', 'Number') else islands.get(n, '1') if f == 'IslandNumber'
                          else 'YES' if f == 'Slack' and n in self.slack else values.get(f, 'NO' if f == 'Slack' else '0')
                          for f in fields] for n in buses])
        return ('', [])

def swing_case():
    gen = gens([('xx', 9, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '1', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[99, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    return gen, load

def test_midrun_balancer_keeps_the_swing_near_zero_and_never_moves_it():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    # Scale load from 200 to 3200 MW with gens fixed: without balancing, the swing would end at +3000 MW.
    targets = load_ramp_targets()
    after_step, log = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
    # The scaling log of every balance is kept in its run.
    assert [r.name for r in log] == [f'midrun{n}' for n in range(1, len(log) + 1)] and all(len(r.log['step_df']) for r in log)
    swing = lambda: area.swing_mw(sa.table('Gen'), [])[0]
    assert len(log) >= 1 and all(abs(r.after) < abs(r.before) for r in log)
    assert abs(swing()) <= area.BALANCE_TRIGGER_MW                      # Balanced as it went.
    assert ('Gen', '99') not in sa.applied                              # The swing unit was never written.
    # Balancing moves gen only: the load reached its own target.
    assert targets['Load'].loc['l1', 'SMW_Target'] == 3200 and abs(targets['Load'].loc['l1', 'SMW'] - 3200) < 1e-6
    area.balance_swing(sa)
    assert abs(swing()) < area.BALANCE_TOL_MW and ('Gen', '99') not in sa.applied

def load_ramp_targets():
    """Load 200 -> 3200 MW with gens fixed: without balancing, the swing would end at +3000 MW."""
    return {'Gen': targets_df(ObjectID=['g1'], BusNum=[1], ID='1', Status='Closed', Status_Target='Closed',
                              MWSetPoint=[100.0], MWSetPoint_Target=[100.0]),
            'Load': targets_df(ObjectID=['l1'], BusNum=[3], ID='1', Status='Closed', Status_Target='Closed',
                               DistStatus='Open', DistStatus_Target='Open', AreaNumber=1,
                               **{c: [v] for c, v in zip(scale.SCALED['Load'], [200.0, 20.0, 0, 0, 0, 0, 0, 0])},
                               **{f'{c}_Target': [v] for c, v in zip(scale.SCALED['Load'], [3200.0, 320.0, 0, 0, 0, 0, 0, 0])})}

def test_swing_band_and_limit_trigger():
    gen, load = swing_case()
    gen.loc[gen['ID'] == 'xx', ['MWMax', 'MWMin']] = [500.0, -300.0]
    saved = area.BALANCE_TRIGGER_MW, area.BALANCE_AT_SWING_LIMITS
    try:
        area.BALANCE_TRIGGER_MW, area.BALANCE_AT_SWING_LIMITS = 1000.0, True
        assert area.swing_band(gen, []) == (-300.0, 500.0)                   # The narrower of the two.
        area.BALANCE_TRIGGER_MW = 0.0
        assert area.swing_band(gen, []) == (-300.0, 500.0)
        area.BALANCE_AT_SWING_LIMITS = False
        assert area.swing_band(gen, []) == (-np.inf, np.inf)                 # Off.

        sa = SwingSimAuto(gen, load)
        after_step, log = area.midrun_balancer(sa)
        targets = load_ramp_targets()
        scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
        assert log == [] and targets['Load'].loc['l1', 'SMW'] == 3200      # Never balanced.

        area.BALANCE_AT_SWING_LIMITS = True                                  # Balance at the swing's MWMax (500 MW).
        sa = SwingSimAuto(gen, load)
        after_step, log = area.midrun_balancer(sa)
        scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
        assert len(log) >= 1 and all(r.before > 500 for r in log)
        assert area.swing_mw(sa.table('Gen'), [])[0] <= 500
    finally:
        area.BALANCE_TRIGGER_MW, area.BALANCE_AT_SWING_LIMITS = saved

def test_giant_swing_uses_the_mw_limit_settings():
    sa = FakeSimAuto()
    saved = prep.SWING_MW_MAX, prep.SWING_MW_MIN
    try:
        prep.SWING_MW_MAX, prep.SWING_MW_MIN = 800.0, -100.0
        prep.create_giant_swing(sa, 1)
    finally:
        prep.SWING_MW_MAX, prep.SWING_MW_MIN = saved
    row = sa.applied[('Gen', '1')]
    assert (row['MWMax'], row['MWMin'], row['ID']) == ('800.0', '-100.0', 'xx')

def test_rescale_targets():
    targets = {'Gen': targets_df(ObjectID=['g1', 'g2'], MWSetPoint=[50.0, 10.0], MWSetPoint_Start=[0.0, 10.0], MWSetPoint_Target=[100.0, 10.0]),
               'Load': targets_df(ObjectID=['l1'], **{c: [10.0] for c in ['SMW', 'SMvar', 'IMW', 'IMvar', 'ZMW', 'ZMvar']},
                                  **{f'{c}_Start': [10.0] for c in ['SMW', 'SMvar', 'IMW', 'IMvar', 'ZMW', 'ZMvar']},
                                  **{f'{c}_Target': [20.0] for c in ['SMW', 'SMvar', 'IMW', 'IMvar', 'ZMW', 'ZMvar']})}
    moved = {'Gen': targets_df(ObjectID=['g1', 'gx'], MWSetPoint=[60.0, 5.0], MWSetPoint_Seed=[50.0, 4.0], MWMax=[500.0, 500.0])}
    area.rescale_targets(targets, moved)
    assert targets['Gen'].loc['g1', ['MWSetPoint', 'MWSetPoint_Start', 'MWSetPoint_Target']].tolist() == [60.0, 0.0, 120.0]
    assert targets['Gen'].loc['g2', 'MWSetPoint_Target'] == 10.0                    # Not moved.
    assert targets['Load'].loc['l1', ['SMvar', 'SMvar_Start', 'SMvar_Target']].tolist() == [10.0, 10.0, 20.0]   # Balances move gens only.

# ------------------ review fixes ------------------
@contextmanager
def saved_settings():
    """Module settings (upper-case globals, and pw's print flags) as they are now, put back afterwards."""
    saved = module_settings()
    prints = pw.print_opened, pw.print_solve_errors
    try:
        yield
    finally:
        restore_settings(saved)
        pw.print_opened, pw.print_solve_errors = prints

def module_settings():
    return {(m, k): v for m in settings.MODULES.values() for k, v in vars(m).items() if k.isupper()}

def restore_settings(saved):
    for (m, k), v in saved.items():
        setattr(m, k, v)

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

def test_rescale_caps_gen_targets_at_mwmax():
    df = targets_df(ObjectID=['a', 'b', 'c'], MWSetPoint=[50.0, 50.0, 50.0], MWSetPoint_Start=[0.0, 0.0, 0.0], MWSetPoint_Target=[100.0, 100.0, 130.0])
    scale.rescale(df, pd.Series({'a': 1.2, 'b': 1.2, 'c': 1.2}), ['MWSetPoint'], cap=pd.Series({'a': 110.0, 'b': 1000.0, 'c': 120.0}))
    assert df['MWSetPoint_Target'].round(6).tolist() == [110.0, 120.0, 130.0]   # c was already above its MWMax: not raised.
    assert df['MWSetPoint'].round(6).tolist() == [60.0] * 3                     # Current values are not capped.

def test_midrun_balance_keeps_gens_within_mwmax():
    # Solar s1 (MWMax 100) has 150 MW to place in the EPC: planned at MWMax, 50 MW unmet, which loads the swing.
    gen = gens([('xx', 9, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('s1', 1, '1', 'Closed', 'PV', 10.0, 100.0, 0.0),
                ('t1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0)]).assign(BusNum=[99, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 110.0, 0.0, 0.0, 10.0)]).assign(BusNum=3, NomkV=230.0, BusName='b', DistStatus='Open', DistMWInput=0.0, DistMvarInput=0.0)
    epc_gen = gens([('e1', 1, '1', 'Closed', 'PV', 150.0, 200.0, 0.0), ('e2', 1, '2', 'Closed', 'GT', 450.0, 1000.0, 0.0)]).assign(BusNum=[1, 2])
    epc_load = loads([('e1', 1, 'Closed', 600.0, 0.0, 0.0, 60.0)])
    sa = SwingSimAuto(gen, load)
    with saved_settings():
        area.BALANCE_TRIGGER_MW = 20.0
        area.MAKEUP = False             # Else the plan puts s1's 50 MW on t1 and the swing never moves.
        targets, _ = area.build_targets({'Gen': epc_gen, 'Load': epc_load}, area.read_case(sa), swing_buses=[99])
        after_step, log = area.midrun_balancer(sa)
        scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
        area.balance_swing(sa)
    assert len(log) >= 1
    final = sa.table('Gen').set_index('ObjectID')
    assert final.loc['s1', 'MWSetPoint'] <= final.loc['s1', 'MWMax'] + 1e-9
    assert (targets['Gen']['MWSetPoint_Target'] <= targets['Gen']['MWMax'] + 1e-9).all()

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
    # Storage keeps its seed status at 0, and a unit with no capacity is left alone.
    rows = gens([('b1', 1, '1', 'Closed', 'BA', 5.0, 50.0, -50.0), ('c1', 1, '2', 'Closed', 'SC', 0.0, 0.0, 0.0)])
    assert area.commit(rows, 0.0, 'Thermal', set()).tolist() == [False, True]

def test_decommit_skips_a_unit_which_cannot_go_and_tries_the_next():
    # Need 520 MW. A (500) is first in line (offline in the EPC) but removing it leaves 100. B (50) can go; C can't.
    rows = gens([('A', 1, '1', 'Closed', 'GT', 1.0, 500.0, 0.0), ('B', 1, '2', 'Closed', 'GT', 1.0, 50.0, 0.0), ('C', 1, '3', 'Closed', 'GT', 1.0, 50.0, 0.0)])
    assert area.commit(rows, 468.0, 'Thermal', {'B', 'C'}).tolist() == [True, False, True]

def test_balance_df_counts_the_whole_load_change():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load.assign(IMW=100.0))      # 300 MW of load against 200 MW of gen: the swing is at +100.
    first = scale.runs_sheet(area.balance_swing(sa)).iloc[0]
    assert first['Run'] == 'balance1' and abs(first['Swing MW Before'] - 100) < 1e-9
    assert first['Load MW Moved'] == 0 and abs(first['Gen MW Moved'] - 100) < 1e-6    # All onto gen, I part included.

def test_iterate_raises_when_the_case_does_not_solve():
    class Diverges(FakeSimAuto):
        def diverged(self):
            return True
    targets, _ = scaling_targets(gen_status='Closed')
    try:
        scale.iterate_to_gen_load_targets(Diverges(), targets)
        assert False
    except pw.SimAutoError:
        pass

def test_test_gen_targets_only_writes_included_rows():
    class Recording(FakeSimAuto):
        writes = []
        def OpenCase(self, fp):
            return ('',)
        def CloseCase(self):
            return ('',)
        def ChangeParametersMultipleElementRect(self, table, fields, rows):
            self.writes += [row[0] for row in rows]
            return super().ChangeParametersMultipleElementRect(table, fields, rows)
    sa = Recording()
    gen = targets_df(ObjectID=['g1', 'sw'], BusNum=[1, 9], ID=['1', 'xx'], Status=['Closed', 'Closed'], Status_Target='Closed',
                     MWSetPoint=[10.0, 0.0], MWSetPoint_Target=[20.0, 0.0])
    scale.exclude(gen, gen.index == 'sw', 'Swing Unit')
    saved, pw.connect = pw.connect, lambda: sa
    try:
        with tempfile.TemporaryDirectory() as d:
            fp = Path(d) / 'x.pwb'
            fp.write_text('')
            out = scale.test_gen_targets(fp, gen)
    finally:
        pw.connect = saved
    assert sa.writes == ['1'] and out['Success'].tolist() == [True, True]

def test_topology_targets_never_scale_the_swing():
    common = dict(BusName='A', NomkV=230.0)
    case_gen = pd.DataFrame({'ObjectID': ['g1', 'g2', 'sw'], 'BusNum': [1, 2, 99], 'ID': ['1', '1', 'xx'], 'Status': 'Closed', 'MWSetPoint': [10.0, 5.0, 300.0], **common})
    gv_gen = pd.DataFrame({'ObjectID': ['g1', 'g2'], 'BusNum': [1, 2], 'ID': '1', 'Status': ['Closed', 'Open'], 'MWSetPoint': [20.0, 0.0], **common})
    load = pd.DataFrame({'ObjectID': ['l1'], 'BusNum': 1, 'ID': '1', 'Status': 'Closed', 'SMW': 50.0, 'SMvar': 5.0, 'IMW': 0.0, 'IMvar': 0.0, 'ZMW': 0.0, 'ZMvar': 0.0, 'DistStatus': 'Closed', 'DistMWInput': 3.0, 'DistMvarInput': 1.0, **common})
    gen = topo.build_targets({'Gen': gv_gen, 'Load': load}, {'Gen': case_gen, 'Load': load}, swing_buses=[2])['Gen']
    # The giant swing, and the gen on the slack bus (bus 2), are excluded and kept Closed even though the EPC has g2 Open.
    assert gen['ExclusionReason'].tolist() == ['', 'Swing Unit', 'Swing Unit']
    assert gen['Status_Target'].tolist() == ['Closed', 'Closed', 'Closed']
    assert topo.build_targets({'Gen': gv_gen, 'Load': load}, {'Gen': case_gen, 'Load': load})['Gen']['ExclusionReason'].tolist() == ['', '', 'Swing Unit']

class CaseSimAuto(SwingSimAuto):
    """SwingSimAuto which also opens and saves cases: `cases` maps a file name to its (gen, load) tables."""
    def __init__(self, cases, slack=()):
        super().__init__(*next(iter(cases.values())), slack=slack)
        self.cases, self.saved = cases, {}

    def OpenCase(self, fp):
        gen, load = self.cases[Path(fp).name]
        self.base = {'Gen': gen.set_index('BusNum', drop=False), 'Load': load.set_index('BusNum', drop=False)}
        self.applied = {}
        return ('',)

    def SaveCase(self, fp, case_format, overwrite):
        self.saved[Path(fp).name] = {'Gen': self.table('Gen').set_index('ObjectID'), 'Load': self.table('Load').set_index('ObjectID')}
        Path(fp).write_text('')
        return ('',)

    def CloseCase(self):
        return ('',)

def load_step_02(folder: Path, mode: str):
    """'02 Scale Hours.py' (not importable by name), pointed at `folder`. SimAuto isn't connected: no win32com needed."""
    spec = importlib.util.spec_from_file_location('step_02', Path(__file__).resolve().parents[1] / '02 Scale Hours.py')
    step = importlib.util.module_from_spec(spec)
    saved = module_settings()
    spec.loader.exec_module(step)           # Loads settings.toml: the owner's choices must not reach the tests.
    restore_settings(saved)
    step.run = dict(settings.RUN_DEFAULTS, mode=mode)
    step.gv_dir, step.out_dir = folder / 'HourEPCs', folder / 'Output'
    step.pw_fp, step.pvqv_fp, step.dummy_bus_fp = folder / 'Prepared.pwb', folder / 'pvqv.csv', folder / 'DummyBus.aux'
    for d in (step.gv_dir, step.out_dir):
        d.mkdir()
    for fp in (step.pw_fp, step.gv_dir / 'h1.epc'):
        fp.write_text('')
    return step

def two_cases():
    """The PWB: giant swing xx (bus 99), g1 thermal, h1 hydro, load l1. The EPC: other keys in area mode; the same in topology mode."""
    gen = gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0),
                ('h1', 1, '1', 'Closed', 'HY', 100.0, 1000.0, 0.0)]).assign(BusNum=[99, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    return gen, load

def test_step_02_area_mode_end_to_end():
    gen, load = two_cases()
    epc_gen = gens([('e1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0), ('e2', 1, '2', 'Closed', 'HY', 50.0, 1000.0, 0.0)]).assign(BusNum=[1, 2], MW=0.0, NomkV=230.0, BusName='b')
    epc_load = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        assert (step.out_dir / 'h1.pwb').exists()
        sheets = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name=None)
        assert set(sheets) == {'areas', 'exceptions', 'runs'}
        assert sheets['runs']['Run'].tolist()[:2] == ['prebalance', 'main']        # Targets 100 MW short of load: balanced first.
        assert sorted(fp.name for fp in step.out_dir.glob('*.xlsx')) == ['h1_Log.xlsx']
    out = sa.saved['h1.pwb']
    assert ('Gen', '99') not in sa.applied                                      # The swing was never written.
    assert abs(out['Gen'].loc['xx', 'MW']) < area.BALANCE_TOL_MW                # Balanced to ~0 MW.
    assert (out['Load']['DistStatus'] == 'Open').all() and (out['Load']['Status'] == 'Closed').all()
    areas = sheets['areas'].set_index('Group')
    assert areas['Target MW'].to_dict() == {'Hydro': 50.0, 'Load': 300.0, 'Thermal': 150.0}
    assert areas['Final MW'].notna().all()

def test_step_02_topology_mode_end_to_end():
    gen, load = two_cases()
    epc_gen = gen.assign(MWSetPoint=[0.0, 150.0, 50.0]).iloc[1:]
    epc_load = load.assign(SMW=300.0)
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'topology')
        step.run['balance_swing'] = False      # Exact EPC values.
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99, 2])   # h1 is on a slack bus.
        step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        log = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name=None)
        assert set(log) == {'exceptions', 'runs'} and list(log['runs']['Run']) == ['main']
        assert sorted(fp.name for fp in step.out_dir.glob('*.xlsx')) == ['h1_Log.xlsx']
    out = sa.saved['h1.pwb']
    assert out['Gen']['MWSetPoint'].round(6).to_dict() == {'xx': 0.0, 'g1': 150.0, 'h1': 100.0}   # h1: left alone.
    assert out['Load']['SMW'].round(6).tolist() == [300.0]
    assert ('Gen', '99') not in sa.applied and ('Gen', '2') not in sa.applied   # The swing was never written.

def test_step_02_topology_mode_aims_its_own_slack_at_its_epc_value():
    # No giant swing. s1 (bus 2, MWMin blank = unlimited below) is the main island's slack; s2 (bus 5, island 2) is another
    # island's slack. Both are in the EPC, but only s1 has a goal: its EPC MW, 300 (the blank MWMin must not clip it to MWMax).
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, np.nan), ('s2', 1, '2', 'Closed', 'GT', 0.0, 1000.0, 0.0),
                ('g1', 1, '3', 'Closed', 'GT', 100.0, 5000.0, 0.0)]).assign(BusNum=[2, 5, 1], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    epc_gen, epc_load = gen.assign(MWSetPoint=[300.0, 400.0, 100.0]), load.assign(SMW=500.0)
    class TwoIslands(CaseSimAuto):
        def table(self, name):
            df = super().table(name)
            if name == 'Gen':
                df.loc[df['BusNum'] == 5, 'MW'] = 0.0         # Island 2 balances itself: nothing reaches the main island's swing.
            return df
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'topology')
        step.run['balance_swing'] = True
        area.BALANCE_TRIGGER_MW = 100.0
        sa = TwoIslands({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[2, 5])
        sa.islands = {5: '2'}
        real, got = step.topo.compute_pw_targets, {}
        def spy(*args):
            got['targets'], got['reports'] = real(*args)
            return got['targets'], got['reports']
        step.topo.compute_pw_targets = spy
        try:
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        finally:
            step.topo.compute_pw_targets = real
        runs = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name='runs')
    swing = got['reports']['swing']
    assert swing.index.tolist() == ['s1'] and swing['EPC MW'].tolist() == [300.0], swing
    assert abs(runs['Swing MW After'].iloc[-1] - 300.0) <= area.BALANCE_TOL_MW, runs
    # The pre-balance counts the slack's planned +200 MW: the other gens take only the rest of the load's +300 (105 with dist gen)...
    assert runs['Run'].iloc[0] == 'prebalance' and abs(runs['Gen MW Moved'].iloc[0] - 105.0) < 1e-6, runs
    # ...and the midrun balancer expects the slack to ride up to 300 MW with the scaling (trigger 100), so it never balances.
    assert not runs['Run'].str.startswith('midrun').any(), runs

def test_slack_bus_gens_are_the_swing_when_there_is_no_giant_swing():
    # Bus 9 is the slack bus: its unit carries the imbalance, and neither balance writes to it.
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    sa = SwingSimAuto(gen, load, slack=[9])
    assert prep.slack_buses(sa) == [9]
    after_step, log = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
    assert len(log) >= 1 and ('Gen', '9') not in sa.applied
    area.balance_swing(sa)
    # s1's key isn't in the EPC: it only has to end within its limits (0 to 1000 MW), not at 0.
    assert -area.BALANCE_TOL_MW < area.swing_mw(sa.table('Gen'), [9])[0] < 1000 + area.BALANCE_TOL_MW
    assert ('Gen', '9') not in sa.applied

def test_own_slack_with_wide_limits_is_kept_near_its_plan_while_scaling():
    # s1 can run -10 to +10 GW, but strays of ~2 GW diverge: it is balanced back once it is balance_trigger_mw (1000)
    # from where the plan expects it (here its start, 0 MW: its key isn't in the EPC). Load +3000 MW, no pre-balance.
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 0.0, 1e4, -1e4), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    sa = SwingSimAuto(gen, load, slack=[9])
    after_step, log = area.midrun_balancer(sa)
    seen = []
    def watch(SimAuto, targets, progress):
        seen.append(area.swing_mw(sa.table('Gen'), [9])[0])         # Before any balance: what the step left.
        return after_step(SimAuto, targets, progress)
    scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=watch)
    assert len(log) >= 1 and all(abs(r.after) < 1.0 for r in log)                      # Back to its plan, 0 MW.
    assert max(abs(x) for x in seen) < 1100 and ('Gen', '9') not in sa.applied           # The step is capped too.
    # With an EPC goal, the plan moves from start to goal as scaling goes: at 50 of 100, half way.
    swing = pd.DataFrame({'Seed MW': [0.0], 'EPC MW': [800.0]}, index=['s1'])
    sa = SwingSimAuto(gen, load, slack=[9])
    after_step, log = area.midrun_balancer(sa, swing)
    sa.base['Load']['SMW'] += 2000.0                                                    # Swing jumps to ~2000.
    after_step(sa, {t: df.assign(HeldReason='') for t, df in load_ramp_targets().items()}, 50)
    assert len(log) == 1 and abs(log[0].after - 400.0) < 1.0

def test_own_slack_planned_at_its_limit_does_not_balance_every_step():
    # s1 is at its MWMax (0 MW) and stays planned there; load +3000 MW, no pre-balance. Balancing aims 500 MW inside
    # (half of balance_trigger_mw), so the drift has room: a few balances, not one per 1% step.
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 0.0, 0.0, -1e4), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    sa = SwingSimAuto(gen, load, slack=[9])
    after_step, log = area.midrun_balancer(sa)
    out = scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
    assert 1 <= len(log) <= 10 and len(out['step_df']) <= 25, (len(log), len(out['step_df']))
    assert all(abs(r.after + 500.0) < 1.0 for r in log)                                # 500 MW inside its MWMax.

def test_own_slack_ends_at_its_epc_mw_when_its_key_is_in_the_epc():
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    epc_gen = gens([('s1', 1, '1', 'Closed', 'GT', 300.0, 1000.0, 0.0)]).assign(BusNum=9)
    swing = area.swing_epc(epc_gen, gen.assign(MW=[150.0, 0.0, 0.0]), [9])
    assert swing.to_dict('index') == {'s1': {'Seed MW': 150.0, 'EPC MW': 300.0}}
    assert area.swing_goal(gen, [9], swing) == (300.0, 300.0) and area.swing_goal(gen, [9]) == (0.0, 1000.0)
    giant = pd.concat([gen, gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4)]).assign(BusNum=9)])
    assert area.swing_epc(epc_gen, giant, [9]).empty and area.swing_goal(giant, [9], swing) == (0.0, 0.0)
    sa = SwingSimAuto(gen, load, slack=[9])
    after_step, log = area.midrun_balancer(sa, swing)
    scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
    area.balance_swing(sa, epc=swing)
    assert abs(area.swing_mw(sa.table('Gen'), [9])[0] - 300.0) < area.BALANCE_TOL_MW and ('Gen', '9') not in sa.applied
    # The goal stays within the slack's limits: an EPC MW past MWMax, or a slack Open in the EPC (0) under MWMin 50.
    lim = gen.assign(MWMin=[50.0, 0.0, 0.0])
    assert area.swing_goal(lim, [9], pd.DataFrame({'EPC MW': [1500.0]}, index=['s1'])) == (1000.0, 1000.0)
    assert area.swing_epc(epc_gen.assign(Status='Open'), lim.assign(MW=150.0), [9]).loc['s1', 'EPC MW'] == 50.0
    assert area.swing_goal(lim.assign(MWMax=np.nan), [9]) == (50.0, np.inf)               # A blank limit: unlimited.
    assert area.swing_band(lim.assign(MWMin=2000.0), [9]) == (1000.0, 2000.0)             # Swapped limits: in order.
    # The pre-balance makes up the slack's move too: seed 150 -> EPC 300 leaves 150 MW less for the other gens.
    t = load_ramp_targets()
    t['Load']['SMW_Target'] = 200.0
    t['Gen']['MWSetPoint'], t['Gen']['MWSetPoint_Target'], t['Gen']['MWMax'] = 400.0, 400.0, 1000.0
    area.prebalance_targets(t, swing_change=-150.0)
    assert t['Gen'].loc['g1', 'MWSetPoint_Target'] == 250.0 and t['Load'].loc['l1', 'SMW_Target'] == 200.0

def ledger_run(mutate=None, slack=(99,), makeup=False):
    """Midrun + final balances, a diverging gen (g3, bus 4) and the areas sheet of the result. slack: slack buses (99 has xx)."""
    gen = gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '2', 'Closed', 'HY', 100.0, 5000.0, 0.0), ('g3', 1, '3', 'Closed', 'GT', 100.0, 5000.0, 0.0)]
               ).assign(BusNum=[99, 1, 2, 4], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    epc_gen = gens([('g1', 1, '1', 'Closed', 'GT', 200.0, 1000.0, 0.0), ('g2', 1, '2', 'Closed', 'HY', 150.0, 1000.0, 0.0),
                    ('g3', 1, '3', 'Closed', 'GT', 200.0, 1000.0, 0.0)])
    epc_load = loads([('l1', 1, 'Closed', 3200.0, 0.0, 0.0, 320.0)])
    sa = SwingSimAuto(gen, load, slack=[b for b in slack if b != 99])
    sa.bad = {'4'}                                              # Any write to g3 diverges: it is held back.
    targets, reports = area.build_targets({'Load': epc_load, 'Gen': epc_gen}, {'Load': load, 'Gen': gen}, swing_buses=list(slack))
    after_step, midrun_log = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
    tables, mk = [], []
    if makeup:
        assert area.makeup(sa, targets, tables=mk)
    area.balance_swing(sa, targets, tables=tables)
    assert midrun_log and tables
    return area.areas(area.read_case(sa), reports, targets, tables, swing_buses=list(slack), makeup_tables=mk), targets

def test_areas_ledger_holds_with_midrun_and_final_balances_and_a_diverging_gen():
    out, targets = ledger_run()
    assert out.columns.tolist() == area.AREAS_COLUMNS and out.groupby(['AreaNumber', 'Group']).size().max() == 1
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Make-up MW'] + out['Balance MW']
    assert (abs(out['Final MW'] - ledger) < 1e-6).all(), out
    row = lambda g: out[out['Group'] == g].iloc[0]
    assert row('Thermal')['Not Reached MW'] > 0 and row('Thermal')['Balance MW'] != 0     # g3 held at 100 of 200 MW; g1 balanced.
    assert row('Load')['Balance MW'] == 0 and pd.isna(row('Thermal')['Target Mvar']) and row('Load')['Target Mvar'] == 320
    assert (out['Note'] == '').all() and (out['Online Final'] > 0).all()

def test_areas_before_scaling_has_blank_final_columns():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0)])
    epc = gens([('g1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0), ('s1', 1, '2', 'Closed', 'PV', 40.0, 100.0, 0.0)])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case})
    out = area.areas(None, reports)
    assert out[['Group', 'Target MW', 'Unmet MW']].values.tolist() == [['Solar', 40.0, 40.0], ['Thermal', 150.0, 0.0]]
    assert out['Final MW'].isna().all() and (out['Not Reached MW'] == 0).all() and out.loc[0, 'Note'] == 'no PWB units'

# ------------------ exceptions and runs sheets ------------------
def exceptions_run():
    """ledger_run's case, with the runs collected: main, midrun balances, final balances."""
    gen = gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0),
                ('g2', 1, '2', 'Closed', 'HY', 100.0, 1000.0, 0.0), ('g3', 1, '3', 'Closed', 'GT', 100.0, 1000.0, 0.0)]
               ).assign(BusNum=[99, 1, 2, 4], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    epc_gen = gens([('g1', 1, '1', 'Closed', 'GT', 200.0, 1000.0, 0.0), ('g2', 1, '2', 'Closed', 'HY', 150.0, 1000.0, 0.0),
                    ('g3', 1, '3', 'Closed', 'GT', 200.0, 1000.0, 0.0)])
    epc_load = loads([('l1', 1, 'Closed', 3200.0, 0.0, 0.0, 320.0)])
    sa = SwingSimAuto(gen, load)
    sa.bad = {'4'}                                              # Any write to g3 diverges: held in main, then excluded.
    targets, reports = area.build_targets({'Load': epc_load, 'Gen': epc_gen}, {'Load': load, 'Gen': gen}, swing_buses=[99])
    after_step, midruns = area.midrun_balancer(sa)
    main = scale.Run('main', scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step))
    return sa, targets, [main, *midruns, *area.balance_swing(sa, targets)]

def test_exceptions_and_runs_sheets_have_fixed_columns_and_one_row_per_element_and_run():
    sa, targets, runs = exceptions_run()
    ex, rs = scale.exceptions(runs, targets, area.read_case(sa)), scale.runs_sheet(runs)
    assert ex.columns.tolist() == scale.EXCEPTION_COLUMNS and rs.columns.tolist() == scale.RUN_COLUMNS
    assert not ex.duplicated(['Table', 'ObjectID']).any()
    assert rs['Run'].tolist() == [r.name for r in runs] and rs['Run'].iloc[0] == 'main' and len(rs) >= 3
    assert all(n.startswith(('main', 'midrun', 'balance')) and '_' not in n for n in rs['Run'])
    g3 = ex[ex['ObjectID'] == 'g3'].iloc[0]                     # Excluded in main (still held at the end): one row, Thermal group.
    assert g3['Table'] == 'Gen' and g3['Group'] == 'Thermal' and g3['AreaNumber'] == 1 and g3['Reason'].startswith('main: ')
    assert g3['Final MW'] == 100.0 and g3['Target MW'] == 200.0 and g3['Seed MW'] == 100.0
    assert ex['ObjectID'].tolist() == ['g3'] and rs['Steps'].iloc[0] > 0 and rs['Diverged Steps'].iloc[0] >= 0
    assert rs['Gen MW Moved'].iloc[0] != rs['Gen MW Moved'].iloc[0] and rs['Gen MW Moved'].iloc[1] != 0   # Main blank, balance filled.

def test_exceptions_join_reasons_from_main_and_a_balance_and_name_the_balance():
    sa, targets, runs = exceptions_run()
    bal = runs[-1].moved
    bal['Gen'].loc['g3', 'ExclusionReason'] = 'Voltage'         # Pretend the last balance excluded g3 too (and g1 for the first time).
    bal['Gen'].loc['g1', 'ExclusionReason'] = 'Voltage'
    runs[-1].log = {'gen': bal['Gen'][bal['Gen']['ExclusionReason'] != ''], 'load': bal['Load'].iloc[:0]}
    ex = scale.exceptions(runs, targets, area.read_case(sa)).set_index('ObjectID')
    assert ex.loc['g1', 'Reason'] == f'{runs[-1].name}: Voltage'
    assert ex.loc['g3', 'Reason'] == f'main: {targets["Gen"].loc["g3", "ExclusionReason"]}; {runs[-1].name}: Voltage'

def test_exceptions_list_network_changes_and_target_status_misses():
    log = {'statcom_bus_df': pd.DataFrame({'BusNum': [5, 6], 'Closed': [True, False]}), 'dropped_branch_df': pd.DataFrame({'ObjectID': ['b1']}),
           'gen': targets_df(ObjectID=['sw'], BusNum=[1]).assign(ExclusionReason='Swing Unit')}
    runs = [scale.Run('main', {'dropped_branch_df': pd.DataFrame({'ObjectID': ['b1']}), 'gen': log['gen']}), scale.Run('midrun3', {'statcom_bus_df': log['statcom_bus_df']})]
    targets = {'Gen': targets_df(ObjectID=['sw', 'g1'], BusNum=[1, 2], Status_Target='Closed', MWSetPoint_Seed=[0.0, 1.0], MWSetPoint_Target=[0.0, 5.0],
                                 MWSetPoint=[0.0, 5.0]).assign(ExclusionReason=['Swing Unit', '']),
               'Load': targets_df(ObjectID=[], BusNum=[])}
    final = {'Gen': pd.DataFrame({'ObjectID': ['sw', 'g1'], 'Status': ['Open', 'Open'], 'MWSetPoint': [0.0, 0.0]}), 'Load': pd.DataFrame({'ObjectID': []})}
    ex = scale.exceptions(runs, targets, final)
    assert ex[['Table', 'ObjectID', 'Reason']].values.tolist() == [
        ['Gen', 'g1', 'not on target status'], ['STATCOM', '5 SC', 'midrun3: STATCOM added'],
        ['STATCOM', '6 SC', 'midrun3: STATCOM tried, rolled back (not in the case)'], ['Branch', 'b1', 'main: branch opened (collapse)']]
    assert pd.isna(ex.loc[3, 'AreaNumber']) and ex.loc[0, 'Final Status'] == 'Open' and ex.loc[0, 'Target MW'] == 5.0   # No Swing Unit row.
    assert scale.exceptions([], None, None).columns.tolist() == scale.EXCEPTION_COLUMNS and len(scale.runs_sheet([])) == 0

# ------------------ settings ------------------
def test_repo_settings_file_loads():
    with saved_settings():
        assert settings.load()['mode'] in settings.MODES

def test_settings_reject_a_negative_loss_exponent():
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        fp = Path(d) / 's.toml'
        fp.write_text('[area]\nloss_exponent = -1.0\n')
        try:
            settings.load(fp)
            assert False, 'a negative loss_exponent must be rejected'
        except ValueError as e:
            assert 'loss_exponent' in str(e)

def test_settings_apply_and_reject():
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

def test_trial_refuses_to_nest():
    sa = FakeSimAuto()
    with pw.trial(sa):
        try:
            with pw.trial(sa):
                pass
            assert False
        except RuntimeError:
            pass

# ------------------ Fable review fixes (2026-11-16) ------------------
def test_midrun_balance_aims_inside_the_band():
    # Swing limits 0..800 MW: 0 is the band's edge, so a balance aims at 200 MW (a quarter of the band in), not 0.
    gen, load = swing_case()
    gen.loc[gen['ID'] == 'xx', ['MWMax', 'MWMin']] = [800.0, 0.0]
    with saved_settings():
        area.BALANCE_TRIGGER_MW, area.BALANCE_AT_SWING_LIMITS = 1000.0, True
        sa = SwingSimAuto(gen, load)
        after_step, log = area.midrun_balancer(sa)
        scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
    assert len(log) >= 1 and abs(log[0].after - 200.0) < 1.0

def test_final_balance_keeps_the_hour_when_a_round_fails():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.base['Load'].loc[3, 'SMW'] = 900.0                       # Swing at +700 MW.
    def fails(*args, **kwargs):
        raise pw.SimAutoError('Could not solve the original input case.')
    saved = scale.iterate_to_gen_load_targets
    scale.iterate_to_gen_load_targets = fails
    try:
        runs = area.balance_swing(sa)
    finally:
        scale.iterate_to_gen_load_targets = saved
    assert [r.name for r in runs] == ['balance1'] and runs[0].note.startswith('stopped: Could not solve') and runs[0].after == runs[0].after

def test_step_02_saves_the_case_when_the_log_fails():
    # The final log write fails (open in Excel), or the log cannot be built: either way the saved case is kept.
    gen, load = two_cases()
    for failure in ('write', 'build'):
        with saved_settings(), tempfile.TemporaryDirectory() as d:
            step = load_step_02(Path(d), 'area')
            sa = CaseSimAuto({'h1.epc': (gen, load), 'Prepared.pwb': (gen, load)}, slack=[99])
            real_write, real_areas = pw.write_excel, step.area.areas
            def locked(fp, sheets):
                if fp.name.endswith('_Log.xlsx') and fp.exists():     # Open in Excel at the final write.
                    raise PermissionError(fp)
                return real_write(fp, sheets)
            def broken(final, *args, **kwargs):
                if final is not None:
                    raise ValueError('bug in the log')
                return real_areas(final, *args, **kwargs)
            step.pw.write_excel, step.area.areas = (locked, real_areas) if failure == 'write' else (real_write, broken)
            try:
                step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')   # Must not raise.
            finally:
                step.pw.write_excel, step.area.areas = real_write, real_areas
            assert 'h1.pwb' in sa.saved, failure

def test_adjust_shunts_takes_no_checkpoint_when_nothing_changes():
    class InBand(FakeSimAuto):
        def GetParametersMultipleElementRect(self, table, fields, filter_group):
            if table == 'Shunt':
                row = {'ObjectID': 'Shunt 1 1', 'MvarNom': '50', 'Status': 'Closed', 'Vpu': '1.0'}
                return ('', [[row[f] for f in fields]])
            return super().GetParametersMultipleElementRect(table, fields, filter_group)
    sa = InBand()
    prep.adjust_shunts(sa)
    assert sa.checkpoint is None and sa.applied == {}

def test_midrun_balance_failure_keeps_the_hour_and_its_network_log():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    after_step, log = area.midrun_balancer(sa)
    run = scale.iterate_to_gen_load_targets
    def balance_fails(SimAuto, targets, **kwargs):
        e = pw.SimAutoError('Could not solve the original input case.')
        e.log = {'statcom_bus_df': pd.DataFrame({'BusNum': [7], 'Closed': [True]})}
        raise e
    scale.iterate_to_gen_load_targets = balance_fails     # Only the balances' nested runs look it up by name.
    try:
        targets = load_ramp_targets()
        run(sa, targets, after_step=after_step)
    finally:
        scale.iterate_to_gen_load_targets = run
    assert targets['Load'].loc['l1', 'Progress'] == scale.RESOLUTION                 # The main scaling finished.
    assert log and log[0].note.startswith('stopped: Could not solve')
    assert log[0].log['statcom_bus_df']['BusNum'].tolist() == [7]                   # The STATCOM it added is listed.
    ex = scale.exceptions(log, None, None)
    # Each failed balance listed it (steps are capped now, so the swing is caught again before it runs far).
    assert ex[['Table', 'ObjectID', 'BusNum']].values.tolist() == [['STATCOM', '7 SC', 7]]
    assert ex['Reason'].iloc[0].startswith('midrun1: STATCOM added')

def test_balance_starts_at_the_largest_step():
    log = scale.iterate_to_gen_load_targets(FakeSimAuto(), scaling_targets(gen_status='Closed')[0], initial_step=scale.MAX_STEP)
    assert log['step_df'].iloc[0][['From', 'To']].tolist() == [0, scale.RESOLUTION]

def test_clear_small_islands_error_is_not_fatal():
    class Collapsed(FakeSimAuto):
        def GetParametersMultipleElementRect(self, table, fields, filter_group):
            if table == 'Branch':
                row = {'ObjectID': "Branch '1' '2' '1'", 'Status': 'Closed', 'BranchVpuHigh': '1.0', 'BranchVpuLow': '0.5'}
                return ('', [[row[f] for f in fields]])
            return super().GetParametersMultipleElementRect(table, fields, filter_group)
        def RunScriptCommand(self, command):
            return ('ClearSmallIslands: not allowed in Edit mode',) if command.startswith('ClearSmallIslands') else super().RunScriptCommand(command)
    assert scale.drop_collapsed_sections(Collapsed()) == ["Branch '1' '2' '1'"]

def test_final_balance_leaves_units_the_main_scaling_excluded():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.base['Load'].loc[3, 'SMW'] = 600.0                                          # Swing at +400 MW.
    scaled = {'Gen': targets_df(ObjectID=['g1'], BusNum=[1]).assign(ExclusionReason='PVQV', HeldReason=''),
              'Load': targets_df(ObjectID=['l1'], BusNum=[3]).assign(HeldReason='')}
    area.balance_swing(sa, scaled)
    out = sa.table('Gen').set_index('ObjectID')
    assert out.loc['g1', 'MWSetPoint'] == 100.0 and out.loc['g2', 'MWSetPoint'] > 100.0

def test_midrun_balance_leaves_units_the_main_scaling_excluded():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    after_step, log = area.midrun_balancer(sa)
    targets = load_ramp_targets()
    targets['Gen']['ExclusionReason'] = 'PVQV'                                     # g1: excluded by the main scaling.
    scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
    out = sa.table('Gen').set_index('ObjectID')
    assert log and out.loc['g1', 'MWSetPoint'] == 100.0 and out.loc['g2', 'MWSetPoint'] > 100.0

def test_merge_reports_merges_hour_logs():
    spec = importlib.util.spec_from_file_location('merge_reports', Path(__file__).resolve().parents[1] / '03 Merge Reports.py')
    merge_reports = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(merge_reports)
    with tempfile.TemporaryDirectory() as d:
        out_dir = Path(d) / 'Output'
        out_dir.mkdir()
        areas_a = pd.DataFrame({'AreaNumber': [1, 1], 'Group': ['Load', 'Coal'], 'Unmet MW': [1.0, 0.0], 'Not Reached MW': [0.5, 0.0],
                                'Balance MW': [2.0, 3.0], 'Error MW': [-4.0, 1.0]}).reindex(columns=area.AREAS_COLUMNS)
        exc_a = pd.DataFrame({'Table': ['Gen', 'STATCOM'], 'ObjectID': ['g1', 'SC1']}).reindex(columns=scale.EXCEPTION_COLUMNS)
        runs_a = pd.DataFrame({'Run': ['main', 'balance1'], 'Swing MW After': [10.0, 25.0]}).reindex(columns=scale.RUN_COLUMNS)
        pw.write_excel(out_dir / 'HourA_Log.xlsx', {'areas': areas_a, 'exceptions': exc_a, 'runs': runs_a})
        runs_b = pd.DataFrame({'Run': ['main', 'balance1'], 'Swing MW After': [9.0, 7.0],      # Final balance ran out of rounds.
                               'Note': ['', 'stopped at max rounds']}).reindex(columns=scale.RUN_COLUMNS)
        pw.write_excel(out_dir / 'HourB_Log.xlsx', {'areas': pd.DataFrame(columns=area.AREAS_COLUMNS),
                                                    'exceptions': pd.DataFrame(), 'runs': runs_b})
        # HourC: the plan's log, left by a failed hour (no runs): not a result, but its areas rows are still merged.
        areas_c = pd.DataFrame({'AreaNumber': [2], 'Group': ['Load'], 'Target MW': [50.0]}).reindex(columns=area.AREAS_COLUMNS)
        pw.write_excel(out_dir / 'HourC_Log.xlsx', {'areas': areas_c, 'exceptions': pd.DataFrame(),
                                                    'runs': pd.DataFrame(columns=scale.RUN_COLUMNS)})
        target = merge_reports.merge(out_dir)
        sheets = pd.read_excel(target, sheet_name=None)
    assert set(sheets) == {'hours', 'areas', 'exceptions', 'runs'}, set(sheets)
    for name in ('areas', 'exceptions', 'runs'):
        assert sheets[name].columns[0] == 'Hour', name
    assert len(sheets['hours']) == 3
    hours = sheets['hours'].set_index('Hour')
    assert hours['Status'].to_dict() == {'HourA': 'ok', 'HourB': 'check', 'HourC': 'incomplete'}
    assert sheets['areas'].loc[sheets['areas']['Hour'] == 'HourC', 'Target MW'].tolist() == [50.0]
    assert hours.loc['HourA', 'Final Swing MW'] == 25.0 and hours.loc['HourB', 'Final Swing MW'] == 7.0
    assert hours.loc['HourA', 'Balance MW'] == 3.0
    assert hours.loc['HourA', 'Worst Error MW'] == -4.0
    assert hours.loc['HourA', 'Exceptions Gen/Load'] == 1 and hours.loc['HourA', 'Exceptions STATCOM'] == 1
    assert hours.loc['HourB', 'Exceptions Gen/Load'] == 0

def test_step_02_console_is_one_summary_line_per_hour():
    import contextlib, io
    gen, load = two_cases()
    epc_gen = gens([('e1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0), ('e2', 1, '2', 'Closed', 'HY', 50.0, 1000.0, 0.0)]).assign(BusNum=[1, 2], MW=0.0, NomkV=230.0, BusName='b')
    epc_load = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    buf = io.StringIO()
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        with contextlib.redirect_stdout(buf):
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
    lines = buf.getvalue().splitlines()
    summaries = [l for l in lines if 'swing' in l and 'exceptions' in l]
    assert len(summaries) == 1 and summaries[0].startswith('h1: swing') and 'worst error area' in summaries[0], lines
    assert not any('Opened:' in l for l in lines), lines
    assert any('Main scaling: 0 -> ' in l and 'of 100' in l for l in lines), lines

def test_areas_ledger_holds_with_a_second_slack_bus():
    # Bus 2 (hydro g2) is a slack bus beside the giant swing: left out of the areas, its balance moves must be too.
    out, _ = ledger_run(slack=(99, 2))
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Make-up MW'] + out['Balance MW']
    assert (abs(out['Final MW'] - ledger) < 1e-6).all(), out

def test_step_02_log_without_balancing_has_the_end_swing():
    import contextlib, io
    gen, load = two_cases()
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        step.run['balance_swing'] = False
        sa = CaseSimAuto({'h1.epc': (gen, load), 'Prepared.pwb': (gen, load)}, slack=[99])
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        runs = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name='runs')
    assert runs['Run'].tolist() == ['main'] and pd.notna(runs['Swing MW After'].iloc[-1])
    assert 'swing nan' not in out.getvalue()

def test_step_02_plan_left_by_a_failed_hour_has_nothing_reached_yet():
    gen, load = two_cases()
    epc_gen = gens([('e1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0)]).assign(BusNum=[1], MW=0.0, NomkV=230.0, BusName='b')
    epc_load = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    def fails(*args, **kwargs):
        raise pw.SimAutoError('Could not solve the original input case.')
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        saved = step.scale.iterate_to_gen_load_targets
        step.scale.iterate_to_gen_load_targets = fails
        try:
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
            assert False
        except pw.SimAutoError:
            pass
        finally:
            step.scale.iterate_to_gen_load_targets = saved
        plan = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name='areas')
    assert (plan['Not Reached MW'] == 0).all() and (plan['Balance MW'] == 0).all() and plan['Final MW'].isna().all()

def test_prebalance_targets_moves_gen_only_and_leaves_what_gens_cannot_take_on_the_swing():
    # Load 200 -> 600 MW, gens fixed: +400 MW would land on the swing. g1 can take +50 (MWMax 150); load keeps its target,
    # and the other 350 MW stay on the swing.
    t = load_ramp_targets()
    t['Load']['SMW_Target'] = 600.0
    t['Gen']['MWMax'] = 150.0
    moved = area.prebalance_targets(t)
    assert t['Gen'].loc['g1', 'MWSetPoint_Target'] == 150.0
    assert t['Load'].loc['l1', 'SMW_Target'] == 600.0 and t['Load'].loc['l1', 'SMvar_Target'] == 320.0
    assert area.mw_changes(moved) == {'Gen MW Change': 50.0, 'Load MW Change': 0.0}
    t['Gen']['MWMax'] = 1000.0
    area.prebalance_targets(t)
    assert t['Gen'].loc['g1', 'MWSetPoint_Target'] == 500.0
    assert area.prebalance_targets(t) is None                                     # Balanced now: nothing to move.

def test_prebalanced_targets_need_no_midrun_balance():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    targets = load_ramp_targets()                                                # Without it: +3000 MW onto the swing.
    targets['Gen']['MWMax'] = 1e4                                                # g1 100 -> 3100, l1 stays 3200.
    area.prebalance_targets(targets)
    assert targets['Gen'].loc['g1', 'MWSetPoint_Target'] == 3100.0
    after_step, runs = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
    assert runs == [] and abs(area.swing_mw(sa.table('Gen'), [])[0]) < 1e-6

def test_prebalance_targets_counts_only_what_will_move():
    # l1's dist gen (40 MW, Closed) goes to 0: +40 MW onto the swing. l2's dist gen is Open: its 30 MW never injected.
    # l3 is Open in both: its 100 MW never counts. g_open goes Closed -> Open (MW unchanged): its 50 MW leave: +50.
    # gx is excluded: not counted, not moved. gcap (target 300) is already above its MWMax (200), so it isn't raised.
    # +90 in all: g1 takes it all; loads keep their targets.
    gen = targets_df(ObjectID=['g1', 'gx', 'g_open', 'gcap'], BusNum=[1, 2, 3, 4], ID='1', Status='Closed',
                     Status_Target=['Closed', 'Closed', 'Open', 'Closed'], MWSetPoint=[100.0, 100.0, 50.0, 300.0],
                     MWSetPoint_Target=[100.0, 150.0, 50.0, 300.0], MWMax=[1000.0, 1000.0, 1000.0, 200.0])
    gen.loc['gx', 'ExclusionReason'] = 'PVQV'
    load = targets_df(ObjectID=['l1', 'l2', 'l3'], BusNum=[5, 6, 7], ID='1', Status=['Closed', 'Closed', 'Open'],
                      Status_Target=['Closed', 'Closed', 'Open'], DistStatus=['Closed', 'Open', 'Closed'], DistStatus_Target='Open',
                      **{c: [0.0] * 3 for c in scale.SCALED['Load']}, **{f'{c}_Target': [0.0] * 3 for c in scale.SCALED['Load']})
    load['SMW'], load['SMW_Target'] = 100.0, 100.0
    load['DistMWInput'] = [40.0, 30.0, 20.0]
    area.prebalance_targets({'Gen': gen, 'Load': load})
    assert gen['MWSetPoint_Target'].round(9).to_dict() == {'g1': 190.0, 'gx': 150.0, 'g_open': 50.0, 'gcap': 300.0}
    assert load['SMW_Target'].round(9).tolist() == [100.0, 100.0, 100.0]

def test_prebalance_targets_counts_losses():
    # Targets balanced by themselves; losses fall 40 MW: gens -40, load unchanged.
    t = load_ramp_targets()
    t['Load']['SMW_Target'] = 1100.0
    t['Gen']['MWSetPoint_Target'], t['Gen']['MWMax'] = 1000.0, 2000.0
    assert area.prebalance_targets(t, 0.0) is None
    area.prebalance_targets(t, loss_change=-40.0)
    assert abs(t['Gen'].loc['g1', 'MWSetPoint_Target'] - 960.0) < 1e-9
    assert t['Load'].loc['l1', 'SMW_Target'] == 1100.0

def test_loss_estimate_scales_each_area_with_its_load():
    # Area 1: 200 -> 100 MW of scaled load plus 100 MW non-conforming: 300 -> 200, losses 9 * (2/3)^2 = 4.
    # Area 2: no seed load: keeps its losses. Area 3: no loads in the targets: ratio 1.
    load = targets_df(ObjectID=['l1', 'l2', 'lo'], BusNum=[1, 2, 3], ID='1', AreaNumber=[1, 2, 1], Status=['Closed', 'Closed', 'Open'],
                      Status_Target=['Closed', 'Closed', 'Open'], **{c: [0.0] * 3 for c in LOAD_P_ALL},
                      **{f'{c}_Target': [0.0] * 3 for c in LOAD_P_ALL})
    load['SMW'], load['SMW_Target'] = [200.0, 0.0, 500.0], [100.0, 50.0, 0.0]
    area_load = pd.DataFrame({'AreaNumber': [1, 2], 'Left As Is MW': [100.0, 0.0]})
    est = area.loss_estimate({'Load': load}, area_load, pd.Series({1: 9.0, 2: 3.0, 3: 5.0}))
    assert est['Est. Loss MW'].round(9).to_dict() == {1: 4.0, 2: 3.0, 3: 5.0} and est['Seed Loss MW'].to_dict() == {1: 9.0, 2: 3.0, 3: 5.0}
    with saved_settings():
        area.LOSS_EXPONENT = 0.0
        assert (area.loss_estimate({'Load': load}, area_load, pd.Series({1: 9.0}))['Est. Loss MW'] == 9.0).all()

def test_loss_estimate_uses_positive_load_and_caps_the_ratio():
    # Area 1: 100 MW load beside a -95 MW DC load; the load goes to 150. Net 5 -> 55 would give losses x121.
    # Positive load: 100 -> 150, losses 10 * 1.5^2 = 22.5. Area 2: 10 -> 100 MW: ratio capped at 2, losses x4.
    load = targets_df(ObjectID=['l1', 'dc', 'l2'], BusNum=[1, 2, 3], ID='1', AreaNumber=[1, 1, 2], Status='Closed', Status_Target='Closed',
                      **{c: [0.0] * 3 for c in LOAD_P_ALL}, **{f'{c}_Target': [0.0] * 3 for c in LOAD_P_ALL})
    load['SMW'], load['SMW_Target'] = [100.0, -95.0, 10.0], [150.0, -95.0, 100.0]
    est = area.loss_estimate({'Load': load}, pd.DataFrame({'AreaNumber': [1, 2], 'Left As Is MW': 0.0}), pd.Series({1: 10.0, 2: 1.0}))
    assert est['Est. Loss MW'].round(9).to_dict() == {1: 22.5, 2: 4.0}

def test_areas_shows_final_losses_without_a_plan_estimate():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0)])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': case}, {'Load': loads([]), 'Gen': case})
    out = area.areas(None, reports, final_loss=pd.Series({1: 7.0}))
    row = out[out['Group'] == 'Load'].iloc[0]
    assert len(out) == 2 and row['Final Loss MW'] == 7.0 and pd.isna(row['Seed Loss MW'])

def test_report_gen_load_balance_counts_only_closed_units():
    gen = targets_df(ObjectID=['g1', 'g2'], Status=['Closed', 'Open'], Status_Target=['Open', 'Closed'],
                     MWSetPoint=[100.0, 50.0], MWSetPoint_Target=[100.0, 50.0])
    load = targets_df(ObjectID=['l1', 'l2'], Status=['Closed', 'Open'], Status_Target='Closed', DistStatus='Closed', DistStatus_Target='Open',
                      **{c: [10.0, 10.0] for c in scale.SCALED['Load']}, **{f'{c}_Target': [10.0, 10.0] for c in scale.SCALED['Load']})
    out = scale.report_gen_load_balance({'Gen': gen, 'Load': load})
    assert out['Base Case'].tolist() == [100.0, 10.0, 30.0, 80.0] and out['Target Case'].tolist() == [50.0, 0.0, 60.0, -10.0]

def test_step_02_estimates_losses_and_clears_untouched_dist_gen():
    # Seed losses 10 MW; area 1's load goes 200 -> 300 MW (non-conforming ln's 50 MW stays): (350/250)^2 * 10 = 19.6.
    gen, load = two_cases()
    load = pd.concat([load, loads([('ln', 1, 'Closed', 50.0, 0.0, 0.0, 0.0)]).assign(BusNum=4, NomkV=230.0, BusName='b', AGC='NO')], ignore_index=True)
    epc_gen = gens([('e1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0), ('e2', 1, '2', 'Closed', 'HY', 50.0, 1000.0, 0.0)]).assign(BusNum=[1, 2], MW=0.0, NomkV=230.0, BusName='b')
    epc_load = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        sa.losses, sa.count_dist = {1: 10.0}, True
        moves = []
        real = area.prebalance_targets
        area.prebalance_targets = lambda t, loss_change=0.0, **k: moves.append(loss_change) or real(t, loss_change, **k)
        try:
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        finally:
            area.prebalance_targets = real
        sheets = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name=None)
    assert abs(moves[0] - 9.6) < 1e-9
    out = sa.saved['h1.pwb']
    assert out['Load'].loc['ln', 'DistStatus'] == 'Open' and out['Load'].loc['ln', 'DistMWInput'] == 0.0
    assert out['Load'].loc['ln', 'SMW'] == 50.0                                  # Non-conforming: its load is left as is.
    run = sheets['runs'].set_index('Run').loc['distgen']                         # ln's 5 MW leave: the swing takes them.
    assert abs(run['Swing MW After'] - run['Swing MW Before'] - 5.0) < 1e-6 and 'cleared on 1 loads (5 MW)' in run['Note']
    assert sheets['runs']['Run'].tolist().index('distgen') == sheets['runs']['Run'].tolist().index('main') + 1
    row = sheets['areas'].set_index('Group').loc['Load']
    assert (row['Seed Loss MW'], round(row['Est. Loss MW'], 9), row['Final Loss MW']) == (10.0, 19.6, 10.0)
    assert sheets['areas'].set_index('Group').loc['Thermal', ['Seed Loss MW', 'Est. Loss MW', 'Final Loss MW']].isna().all()

def test_exclude_islands_names_the_island_and_targets_it_off():
    # Bus 1 main island, bus 2 island 3, bus 4 blank, bus 5 not in the case. gx (a second island's slack gen) loses
    # its 'Swing Unit' reason: it is turned off with the rest of its island.
    gen = targets_df(ObjectID=['g1', 'g2', 'g4', 'g5', 'gx', 'gs'], BusNum=[1, 2, 4, 5, 2, 2], ID=['1'] * 5 + ['xx'], Status='Closed',
                     Status_Target='Closed', MWSetPoint=50.0, MWSetPoint_Target=60.0)
    gen.loc['gx', 'ExclusionReason'] = 'Swing Unit'
    load = targets_df(ObjectID=['l1', 'l2'], BusNum=[1, 2], ID='1', Status='Closed', Status_Target='Closed', DistStatus='Open',
                      DistStatus_Target='Open', **{c: 5.0 for c in LOAD_P_ALL}, **{f'{c}_Target': 6.0 for c in LOAD_P_ALL})
    islands = pd.Series([1.0, 3.0, np.nan], index=[1, 2, 4])
    scale.exclude_islands({'Gen': gen, 'Load': load}, islands)
    assert gen.loc['gs', 'ExclusionReason'] == '' and gen.loc['gs', 'Status_Target'] == 'Closed'   # The giant swing stays.
    gen = gen.drop(index='gs')
    assert gen['ExclusionReason'].to_dict() == {'g1': '', 'g2': 'Island 3, not the main island', 'g4': 'No island',
                                                'g5': 'No island', 'gx': 'Island 3, not the main island'}
    assert gen['Status_Target'].tolist() == ['Closed', 'Open', 'Open', 'Open', 'Open']
    assert gen['MWSetPoint_Target'].tolist() == [60.0, 0.0, 0.0, 0.0, 0.0]
    assert load['ExclusionReason'].tolist() == ['', 'Island 3, not the main island']
    assert load.loc['l2', 'Status_Target'] == 'Open' and (load.loc['l2', [f'{c}_Target' for c in LOAD_P_ALL]] == 0).all()

def test_turn_off_islanded_writes_island_units_open_and_keeps_those_which_diverge():
    gen = targets_df(ObjectID=['g1', 'g2', 'g3'], BusNum=[1, 2, 3], ID='1', Status='Closed', Status_Target=['Closed', 'Open', 'Open'],
                     MWSetPoint=50.0, MWSetPoint_Target=[60.0, 0.0, 0.0])
    gen.loc[['g2', 'g3'], 'ExclusionReason'] = 'Island 2, not the main island'
    load = targets_df(ObjectID=['l1'], BusNum=[1], ID='1', Status='Closed', Status_Target='Closed', DistStatus='Open',
                      DistStatus_Target='Open', **{c: 5.0 for c in LOAD_P_ALL}, **{f'{c}_Target': 6.0 for c in LOAD_P_ALL})
    load.loc['l1', 'ExclusionReason'] = 'Island 2, not the main island'
    load.loc['l1', ['Status_Target', *[f'{c}_Target' for c in LOAD_P_ALL]]] = ['Open'] + [0.0] * len(LOAD_P_ALL)
    class Ordered(FakeSimAuto):
        tables = []
        def ChangeParametersMultipleElementRect(self, table, fields, rows):
            self.tables.append(table)
            return super().ChangeParametersMultipleElementRect(table, fields, rows)
    sa = Ordered(bad={'3'})                                                      # Turning g3 off diverges.
    scale.turn_off_islanded(sa, {'Gen': gen, 'Load': load})
    assert sa.tables[0] == 'Load'                                                # Loads go first.
    assert sa.applied[('Gen', '2')]['Status'] == 'Open' and ('Gen', '3') not in sa.applied and ('Gen', '1') not in sa.applied
    assert gen.loc['g2', 'Status'] == 'Open' and gen.loc['g2', 'MWSetPoint'] == 0.0 and gen.loc['g3', 'Status'] == 'Closed'

def test_balance_skips_units_outside_the_main_island():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.islands = {2: '2', 3: ''}                                                 # g2 (bus 2) and l1 (bus 3) are islanded.
    final = area.read_case(sa)
    assert area.balance_skips(sa, final) == {'g2', 'l1'}

def test_step_02_excludes_units_outside_the_main_island_before_the_prebalance():
    # h1 (bus 2) is in island 2: excluded with that reason, turned off before scaling, not counted by the pre-balance.
    # The EPC's 50 MW of Hydro has no main-island unit left to take it: Unmet, and the areas row adds up.
    gen, load = two_cases()
    epc_gen = gens([('e1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0), ('e2', 1, '2', 'Closed', 'HY', 50.0, 1000.0, 0.0)]).assign(BusNum=[1, 2], MW=0.0, NomkV=230.0, BusName='b')
    epc_load = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        sa.islands = {2: '2'}
        seen = []
        real_pb, real_write = area.prebalance_targets, pw.write_excel
        plans = []
        area.prebalance_targets = lambda t, loss_change=0.0, **k: seen.append(t['Gen'].loc['h1', 'ExclusionReason']) or real_pb(t, loss_change, **k)
        step.pw.write_excel = lambda fp, sheets: plans.append({k: v.copy() for k, v in sheets.items()}) or real_write(fp, sheets)
        try:
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        finally:
            area.prebalance_targets, step.pw.write_excel = real_pb, real_write
        sheets = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name=None)
    assert seen == ['Island 2, not the main island']
    assert sa.saved['h1.pwb']['Gen'].loc['h1', 'Status'] == 'Open' and sa.saved['h1.pwb']['Gen'].loc['h1', 'MWSetPoint'] == 0.0
    assert plans[0]['exceptions'].set_index('ObjectID').loc['h1', 'Reason'] == 'main: Island 2, not the main island'   # In the plan too.
    exc = sheets['exceptions'].set_index('ObjectID')
    assert exc.loc['h1', 'Reason'] == 'main: Island 2, not the main island'
    hydro = sheets['areas'].set_index('Group').loc['Hydro']
    assert (hydro['Seed MW'], hydro['Unmet MW'], hydro['Final MW']) == (0.0, 50.0, 0.0)
    assert 'outside main island, turned off: 100.0 MW' in hydro['Note'] and 'unexplained' not in hydro['Note']

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

def test_seed_islands_are_not_checked_when_the_swing_is_not_in_island_1():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load, slack=[1])
    assert scale.seed_islands(sa, [1]).to_dict() == {1: 1, 2: 1, 3: 1, 99: 1}
    sa.islands = {1: '2'}
    assert scale.seed_islands(sa, [1]) is None
    sa.islands = {2: '2'}                                                       # A live second island with its own slack bus.
    assert scale.seed_islands(sa, [1, 2]).to_dict()[2] == 2

def test_balance_skips_no_island_when_no_bus_is_in_island_1():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.islands = {n: '' for n in (1, 2, 3, 99)}
    assert area.balance_skips(sa, area.read_case(sa)) == set()

def test_areas_ledger_holds_when_an_island_unit_stays_on():
    # h2 (island 2) couldn't be turned off: out of Not Reached and Final, and its Note says so.
    case_gen = gens([('h1', 1, '1', 'Closed', 'HY', 100.0, 200.0, 0.0), ('h2', 1, '2', 'Closed', 'HY', 100.0, 200.0, 0.0)]).assign(BusNum=[1, 2])
    epc_gen = gens([('e1', 1, '1', 'Closed', 'HY', 150.0, 200.0, 0.0)])
    targets, reports = area.build_targets({'Gen': epc_gen, 'Load': loads([])}, {'Gen': case_gen, 'Load': loads([])},
                                          island_by_bus=pd.Series([1.0, 2.0], index=[1, 2]))
    g = targets['Gen']
    g.loc['h1', 'MWSetPoint'] = 150.0                                            # h1 reached; h2 still Closed at 100.
    final = {'Gen': case_gen.assign(MWSetPoint=[150.0, 100.0]), 'Load': loads([])}
    row = area.areas(final, reports, targets).set_index('Group').loc['Hydro']
    assert (row['Not Reached MW'], row['Final MW'], row['Error MW']) == (0.0, 150.0, 0.0)
    assert 'unexplained' not in row['Note'] and 'of which 100.0 MW stayed on' in row['Note']

def test_areas_ledger_holds_with_a_makeup_and_the_held_gens_shortfall_lands_on_its_group():
    out, targets = ledger_run(makeup=True)
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Make-up MW'] + out['Balance MW']
    assert (abs(out['Final MW'] - ledger) < 1e-6).all(), out
    assert not out['Note'].str.contains('unexplained').any(), out['Note']
    thermal = out[out['Group'] == 'Thermal'].iloc[0]
    assert abs(thermal['Make-up MW'] - thermal['Not Reached MW']) < 1e-6 and thermal['Make-up MW'] > 0   # g1 took g3's.
    assert (out.loc[out['Group'] != 'Thermal', 'Make-up MW'] == 0).all()
    # The shortfall no longer lands on the swing, so balancing moves 100 MW less over the case.
    before, _ = ledger_run()
    assert abs(before['Balance MW'].abs().sum() - out['Balance MW'].abs().sum() - 100.0) < 1e-6

def test_makeup_targets_stay_in_the_group_and_respect_caps_islands_and_signs():
    # Area 1 Thermal: gx excluded 60 MW short; g1 (MWMax 120) and g2 share it by MW but g1 caps at 120: g1 120, g2 140.
    # gi is islanded (target 0, still at 50): not counted. Area 2 Thermal (g9) untouched. Hydro h1 isn't Thermal.
    # Load area 1: lx excluded 30 MW short; l1 (100) takes it, the negative dc load and l9 (area 2) don't.
    tg = targets_df(ObjectID=['g1', 'g2', 'gx', 'gi', 'g9', 'h1'], AreaNumber=[1, 1, 1, 1, 2, 1], BusNum=[1, 2, 3, 4, 5, 6],
                    ID='1', UnitTypeCode=['GT'] * 5 + ['HY'], FuelTypeCode='', Status='Closed',
                    MWSetPoint=[100.0, 100.0, 40.0, 50.0, 100.0, 100.0], MWSetPoint_Target=[100.0, 100.0, 100.0, 0.0, 100.0, 100.0])
    tg.loc['gx', 'ExclusionReason'], tg.loc['gi', 'ExclusionReason'] = 'Diverged during scaling', 'Island 2, not the main island'
    tl = targets_df(ObjectID=['l1', 'lx', 'dc', 'l9'], AreaNumber=[1, 1, 1, 2], BusNum=[1, 2, 3, 5], ID='1', Status='Closed',
                    **{c: 0.0 for c in LOAD_P_ALL}, **{f'{c}_Target': 0.0 for c in LOAD_P_ALL})
    tl['SMW'], tl['SMW_Target'] = [100.0, 20.0, -50.0, 100.0], [100.0, 50.0, -50.0, 100.0]
    tl.loc['lx', 'ExclusionReason'] = 'Diverged during scaling'
    final = {'Gen': tg.reset_index().assign(MWMax=[120.0, 1000.0, 1000.0, 1000.0, 1000.0, 1000.0]),
             'Load': tl.reset_index().assign(AGC='YES', DistStatus='Open')}
    moved, asked = area.makeup_targets(final, {'Gen': tg, 'Load': tl}, swing_buses=[], skip={'gx', 'gi', 'lx'})
    assert asked == {'Gen': 60.0, 'Load': 30.0}
    assert moved['Gen']['MWSetPoint_Target'].round(9).to_dict() == {'g1': 120.0, 'g2': 140.0}
    assert moved['Load'].index.tolist() == ['l1'] and moved['Load'].loc['l1', 'SMW_Target'] == 130.0
    assert area.makeup_targets(final, {'Gen': tg.assign(ExclusionReason=''), 'Load': tl.assign(ExclusionReason='')}, []) is None
    # g1 and g2 capped at MWMax: Thermal's 60 MW go onto the area's Hydro h1, not area 2's Thermal g9.
    final['Gen']['MWMax'] = [100.0, 100.0, 1000.0, 1000.0, 1000.0, 1000.0]
    moved, _ = area.makeup_targets(final, {'Gen': tg, 'Load': tl}, swing_buses=[], skip={'gx', 'gi', 'lx'})
    assert moved['Gen']['MWSetPoint_Target'].round(9).to_dict() == {'h1': 160.0}

def test_step_02_runs_a_makeup_only_in_area_mode_when_something_is_short():
    gen, load = two_cases()
    epc_gen = gens([('e1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0), ('e2', 1, '2', 'Closed', 'HY', 50.0, 1000.0, 0.0)]).assign(BusNum=[1, 2], MW=0.0, NomkV=230.0, BusName='b')
    epc_load = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        sheets = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name=None)
    assert 'makeup' not in sheets['runs']['Run'].tolist()                       # Nothing short: no make-up run.
    assert 'Make-up MW' in sheets['areas'].columns
    with saved_settings(), tempfile.TemporaryDirectory() as d:                  # g1 (bus 1) diverges: Thermal is short.
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        sa.bad = {'1'}
        step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        runs = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name='runs').set_index('Run')
    assert 'gen asked +129 MW, placed +129 MW' in runs.loc['makeup', 'Note']     # No other Thermal unit: h1 (Hydro, same area) takes it.
    with saved_settings(), tempfile.TemporaryDirectory() as d:                  # A make-up error keeps the hour.
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        sa.bad = {'1'}
        real = step.area.makeup_targets
        step.area.makeup_targets = lambda *a, **k: 1 / 0
        try:
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        finally:
            step.area.makeup_targets = real
        runs = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name='runs').set_index('Run')
    assert 'h1.pwb' in sa.saved and runs.loc['makeup', 'Note'].startswith('failed: ZeroDivisionError')
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'topology')
        assert step.area.MAKEUP                                                  # On, but area mode only.

# ------------------ T2: area.py / scale.py bugs (2026-11-17) ------------------
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

def test_areas_online_final_leaves_out_statcoms():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 200.0, 0.0), ('sc', 1, scale.STATCOM_ID, 'Closed', '', 0.0, 0.0, 0.0)])
    epc = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 200.0, 0.0)])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case.iloc[:1]})
    out = area.areas({'Gen': case, 'Load': loads([]).assign(NomkV=1.0)}, reports, targets)
    assert out[['Group', 'Online Final']].values.tolist() == [['Thermal', 1.0]]

def test_midrun_balance_skips_gens_the_scaling_takes_to_zero():
    # g1 is decommitted (Open, 100 -> 0 MW), g3 stays Closed at a 0 MW target; g2 stays at 100. The balances go onto g2 only,
    # so rescale_targets() has nothing to undo and the end targets keep every MW a balance moved.
    gen = gens([('xx', 9, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0), ('g3', 1, '3', 'Closed', 'GT', 100.0, 5000.0, 0.0)]
               ).assign(BusNum=[99, 1, 2, 4], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    t = load_ramp_targets()
    t['Gen'] = targets_df(ObjectID=['g1', 'g2', 'g3'], BusNum=[1, 2, 4], ID='1', Status='Closed', Status_Target=['Open', 'Closed', 'Closed'],
                          MWSetPoint=[100.0] * 3, MWSetPoint_Target=[0.0, 100.0, 0.0])
    sa = SwingSimAuto(gen, load)
    after_step, log = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, t, after_step=after_step)
    assert log and all(r.moved['Gen'].index.tolist() == ['g2'] for r in log)
    g = t['Gen']
    assert g.loc['g1', 'MWSetPoint_Target'] == 0 and g.loc['g3', 'MWSetPoint_Target'] == 0
    kept = g['MWSetPoint_Target'].sum() - 100.0
    assert kept > 100.0 and abs(kept - sum(r.moved['Gen']['MWSetPoint'].sum() - r.moved['Gen']['MWSetPoint_Seed'].sum() for r in log)) < 1e-6

def own_slack_ramp(lo, hi, trigger, down=False):
    """A 3 GW load ramp with the case's own slack (bus 9, MWMin lo, MWMax hi) and no giant swing: (midrun runs, main steps, end swing, goal)."""
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 100.0, hi, lo), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    sa = SwingSimAuto(gen, load, slack=[9])
    if down:        # Load 3200 -> 200 instead: the slack drifts down.
        sa.base['Load'].loc[3, 'SMW'] = 3200.0
    with saved_settings(), contextlib.redirect_stdout(io.StringIO()):
        area.BALANCE_TRIGGER_MW = trigger
        after_step, log = area.midrun_balancer(sa)
        t = load_ramp_targets()
        if down:
            for c in scale.SCALED['Load']:
                t['Load'][c], t['Load'][f'{c}_Target'] = t['Load'][f'{c}_Target'].copy(), t['Load'][c].copy()
        out = scale.iterate_to_gen_load_targets(sa, t, after_step=after_step)
        area.balance_swing(sa)
    return log, out['step_df'], area.swing_mw(sa.table('Gen'), [9])[0], area.swing_goal(sa.table('Gen'), [9])

def test_midrun_balance_of_an_own_slack_with_no_range_stays_cheap():
    log, steps, swing, (lo, hi) = own_slack_ramp(100.0, 100.0, 1000.0)       # MWMin == MWMax: the band used to have no width.
    assert len(log) <= 10 and len(steps) <= 25 and lo - area.BALANCE_TOL_MW <= swing <= hi + area.BALANCE_TOL_MW, (len(log), len(steps), swing)
    log, steps, swing, (lo, hi) = own_slack_ramp(100.0, 100.0, 1000.0, down=True)    # The lower half of the band, on a falling load.
    assert len(log) <= 10 and len(steps) <= 25 and lo - area.BALANCE_TOL_MW <= swing <= hi + area.BALANCE_TOL_MW, (len(log), len(steps), swing)

def test_midrun_balance_of_an_own_slack_is_off_with_no_trigger_and_a_narrow_range():
    # Trigger 0: the band is the slack's MWMin to MWMax; a range under 100 MW is not balanced midrun (the final balance still runs).
    # (The plan's example 50..150 is exactly 100 MW wide: 90 MW here, see the T2 report.)
    log, steps, swing, (lo, hi) = own_slack_ramp(50.0, 140.0, 0.0)
    assert log == [] and lo - area.BALANCE_TOL_MW <= swing <= hi + area.BALANCE_TOL_MW
    log, steps, swing, _ = own_slack_ramp(0.0, 1000.0, 0.0)                  # A wide range: balanced at its limits, not every step.
    assert 1 <= len(log) <= 5, len(log)
    log, steps, swing, _ = own_slack_ramp(50.0, 150.0, 0.0)                  # Exactly MIN_OWN_RANGE_MW wide: still balanced (strict <).
    assert len(log) >= 1

def test_a_unit_on_a_bus_out_of_band_from_the_start_is_held_before_scaling_begins():
    gen = targets_df(ObjectID=['g1', 'g2'], BusNum=[1, 2], ID='1', Status='Closed', Status_Target='Closed',
                     MWSetPoint=[10.0, 10.0], MWSetPoint_Target=[20.0, 20.0])
    gen['HeldReason'] = ''
    scale.hold_voltage({'Gen': gen}, pd.Series({1: 1.0, 2: 0.85}))          # No Progress column yet (before mark_start()).
    assert gen['HeldReason'].tolist() == ['', 'Voltage']

def test_areas_keeps_a_seed_gen_whose_id_is_the_statcom_id():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 200.0, 0.0), ('g2', 1, scale.STATCOM_ID, 'Closed', 'GT', 50.0, 100.0, 0.0),
                 ('new', 1, scale.STATCOM_ID, 'Closed', '', 0.0, 0.0, 0.0)])
    epc = case.iloc[:2]
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case.iloc[:2]})
    out = area.areas({'Gen': case, 'Load': loads([]).assign(NomkV=1.0)}, reports, targets).set_index('Group')
    assert out.loc['Thermal', 'Final MW'] == 150.0 and out.loc['Thermal', 'Online Final'] == 2.0   # 'new' is a STATCOM added later.
    assert not out['Note'].str.contains('unexplained').any(), out['Note']

class LowAtEndSimAuto(FakeSimAuto):
    """Bus 2 sags to 0.85 pu once gen g2 has been written at 0 MW (its target): it reached the target, then the voltage dropped."""
    def GetParametersMultipleElementRect(self, table, fields, filter_group):
        if table != 'Bus':
            return ('', [])
        low = any(t == 'Gen' and k == '2' and float(r.get('MWSetPoint', 1)) == 0.0 for (t, k), r in self.applied.items())
        v = lambda n: '0.85' if (n == 2 and low) else '1.0'
        vals = {'IslandNumber': '1', 'BusIsStarBus:1': 'NO', 'BusNomVolt': '230'}
        return ('', [[str(n) if f in ('Busnum', 'Number') else v(n) if f == 'Vpu' else vals.get(f, '0') for f in fields] for n in (1, 2)])

def test_a_unit_at_its_target_is_not_held_for_voltage_and_gets_its_status():
    gen = targets_df(ObjectID=['g1', 'g2'], BusNum=[1, 2], ID='1', Status='Closed', Status_Target=['Closed', 'Open'],
                     MWSetPoint=[10.0, 50.0], MWSetPoint_Target=[20.0, 0.0])
    load = targets_df(ObjectID=['l1'], BusNum=[1], ID='1', Status='Closed', Status_Target='Closed', DistStatus='Open', DistStatus_Target='Open',
                      **{c: [5.0] for c in scale.SCALED['Load']}, **{f'{c}_Target': [5.0] for c in scale.SCALED['Load']})
    t = {'Gen': gen, 'Load': load}
    with contextlib.redirect_stdout(io.StringIO()):
        scale.iterate_to_gen_load_targets(LowAtEndSimAuto(), t)
    g = t['Gen']
    assert g.loc['g2', 'Status'] == 'Open' and g.loc['g2', 'ExclusionReason'] == '' and g.loc['g2', 'Progress'] == scale.RESOLUTION
    assert 'g2' not in set(scale.exceptions([], t)['ObjectID'])

# ------------------ robustness (settings, merge, step 02, prep) ------------------
def test_settings_allowed_keys_are_the_shipped_files_keys():
    import tomllib
    text = (Path(__file__).resolve().parents[1] / 'settings.toml').read_text(encoding='utf-8')
    data = tomllib.loads(text)
    assert set(data['run']) == set(settings.RUN_DEFAULTS)
    assert set(data) - {'run'} == set(settings.SETTINGS)
    for section, keys in settings.SETTINGS.items():
        shipped = set(data[section])
        if section == 'scale':
            assert '# max_workers =' in text     # Commented out in the file: the default is the CPU count.
            shipped.add('max_workers')
        assert shipped == keys, (section, shipped ^ keys)
        assert all(hasattr(settings.MODULES[section], k.upper()) for k in keys), section

def test_settings_reject_values_that_hang_or_disable_the_engine_and_change_nothing():
    first = '[topology]\ntap_threshold = 0.3\n'      # Valid, and read before the bad value: must not stick either.
    bad = ['[scale]\nmax_step = 0\n', '[scale]\ninitial_step = 0\n', '[scale]\ninitial_step = -5\n',
           '[scale]\ninitial_step = 50\nmax_step = 25\n', '[scale]\nmax_step = 101\n', '[scale]\ninitial_step = 2.5\n',
           '[scale]\nstep_solves = 0\n', '[scale]\nfinal_max_rounds = 0\n',
           '[scale]\nexclude_vmin = 1.2\n', '[scale]\nexclude_vmin = 1.12\n', '[scale]\nexclude_vmax = 0.8\n',
           '[scale]\nmax_workers = 0\n', '[scale]\nmax_workers = "eight"\n', '[scale]\nmax_workers = 2.5\n',
           '[pw]\nmva_mismatch_threshold = 0\n', '[area]\ncommit_max_cf = 0\n', '[area]\ncommit_max_cf = 1.5\n',
           '[area]\nbalance_max_rounds = -1\n', '[area]\nbalance_trigger_mw = -1.0\n', '[area]\nbalance_tol_mw = 0\n',
           '[area]\nloss_exponent = 3.0\nbalance_tol_mw = 0\n',
           '[prep]\nswing_mw_min = 50000.0\n', '[prep]\nswing_mw_max = -30000.0\n', '[prep]\nswing_mw_max = -20000.0\n']
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        fp = Path(d) / 's.toml'
        before = module_settings()
        for text in bad:
            fp.write_text(first + text)
            try:
                settings.load(fp)
                assert False, text
            except ValueError:
                pass
            assert module_settings() == before, text         # A rejected file changes nothing (LOSS_EXPONENT included).
        # The edges which are allowed.
        for text in ['[scale]\ninitial_step = 1\nmax_step = 1\nstep_solves = 1\nfinal_max_rounds = 1\nmax_workers = 1\n',
                     '[scale]\ninitial_step = 100\nmax_step = 100\n', '[area]\nbalance_max_rounds = 0\nbalance_trigger_mw = 0\ncommit_max_cf = 1.0\n',
                     '[prep]\nswing_mw_max = 1.0\nswing_mw_min = 0.0\n']:
            fp.write_text(text)
            settings.load(fp)

def test_settings_reject_names_which_are_code_not_settings():
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        fp = Path(d) / 's.toml'
        for text in ['[area]\nload_p = ["SMW"]\n', '[prep]\nswing_id = "yy"\n', '[scale]\nstatcom_id = "x"\n',
                     '[scale]\nmain_island = 2\n', '[scale]\nresolution = 10\n', '[area]\nstatus = {}\n']:
            fp.write_text(text)
            try:
                settings.load(fp)
                assert False, text
            except ValueError as e:
                assert 'unknown setting' in str(e), e

def test_loading_step_02_leaves_the_owners_settings_out_of_the_tests():
    # The step imports settings.toml. Whatever it holds (here: a file which turns on balancing off and commit off), the
    # tests start from the code defaults and the default [run] switches, and the print flags are put back.
    with tempfile.TemporaryDirectory() as d:
        fp = Path(d) / 'owner.toml'
        fp.write_text('[run]\nbalance_swing = false\npretest_gens = true\n[area]\ncommit = false\nmakeup = false\n')
        real = settings.load.__defaults__
        original = pw.print_opened, pw.print_solve_errors
        pw.print_opened = pw.print_solve_errors = True       # The step turns both off at import.
        flags = True, True
        with saved_settings():
            before = module_settings()
            settings.load.__defaults__ = (fp,)
            try:
                step = load_step_02(Path(d), 'area')
            finally:
                settings.load.__defaults__ = real
            assert module_settings() == before and area.COMMIT is True and area.MAKEUP is True
            assert step.run == dict(settings.RUN_DEFAULTS, mode='area')
        assert (pw.print_opened, pw.print_solve_errors) == flags
        pw.print_opened, pw.print_solve_errors = original

def test_step_02_stops_clearly_when_topology_mode_has_no_dummy_bus_file():
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        step.check_inputs()                                  # Area mode needs neither file.
        step.run['mode'] = 'topology'
        try:
            step.check_inputs()
            assert False
        except SystemExit as e:
            assert 'DummyBus.aux' in str(e) and 'topology mode' in str(e)
        step.dummy_bus_fp.write_text('')
        step.check_inputs()

def test_giant_swing_on_a_bus_not_in_the_case_is_a_clear_error():
    sa = FakeSimAuto()
    try:
        prep.create_giant_swing(sa, 424242)
        assert False
    except ValueError as e:
        assert '424242' in str(e) and 'not in the case' in str(e)
    assert not sa.applied

def step_02_hour_with_a_failure(fail_balance=False, fail_distgen=False):
    """Runs one area-mode hour in step 02. A non-conforming load (l2, not in the EPC) lets the test push the swing off by
    50 MW after the main run and again after balance round 1, outside what the ledger counts. fail_balance: the read at the
    start of balance round 2 raises. fail_distgen: the dist gen cleanup raises. Returns (runs sheet, areas sheet, case out)."""
    gen, load = two_cases()
    load = pd.concat([load, loads([('l2', 1, 'Closed', 30.0, 0.0, 0.0, 3.0)]).assign(BusNum=4, NomkV=230.0, BusName='b', AGC='NO')], ignore_index=True)
    epc_gen = gens([('e1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0), ('e2', 1, '2', 'Closed', 'HY', 50.0, 1000.0, 0.0)]).assign(BusNum=[1, 2], MW=0.0, NomkV=230.0, BusName='b')
    epc_load = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    calls, armed = [], []
    with saved_settings(), tempfile.TemporaryDirectory() as d, contextlib.redirect_stdout(io.StringIO()):
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        real_iterate, real_read, real_cleanup = scale.iterate_to_gen_load_targets, area.read_case, area.dist_gen_cleanup
        def iterate(*args, **kwargs):
            out = real_iterate(*args, **kwargs)
            calls.append(kwargs.get('label'))
            if fail_balance and len(calls) <= 2:
                sa.base['Load'].loc[4, 'SMW'] += 50.0
                if len(calls) == 2:
                    armed.append(1)
            return out
        def read_case(*args, **kwargs):
            if armed:
                armed.clear()
                raise RuntimeError('boom')
            return real_read(*args, **kwargs)
        def cleanup(*args, **kwargs):
            raise RuntimeError('no dist gen')
        scale.iterate_to_gen_load_targets, area.read_case = iterate, read_case
        if fail_distgen:
            area.dist_gen_cleanup = cleanup
        try:
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        finally:
            scale.iterate_to_gen_load_targets, area.read_case, area.dist_gen_cleanup = real_iterate, real_read, real_cleanup
        assert (step.out_dir / 'h1.pwb').exists()
        sheets = pd.read_excel(step.out_dir / 'h1_Log.xlsx', sheet_name=None, keep_default_na=False)
    if fail_balance:
        assert calls[:2] == [None, 'Final balance 1'], calls         # Round 1 ran, then round 2's read raised.
    return sheets['runs'], sheets['areas'], sa.saved['h1.pwb']

def test_step_02_keeps_the_hour_when_the_final_balance_raises():
    runs, areas, out = step_02_hour_with_a_failure(fail_balance=True)
    assert runs['Run'].tolist()[-2:] == ['balance1', 'balance2'], runs                 # Round 1 is kept...
    assert pd.to_numeric(runs['Swing MW After'], errors='coerce').notna().iloc[-2], runs  # ...with its After.
    assert runs['Note'].iloc[-1].startswith("failed: RuntimeError('boom')"), runs
    assert not areas['Note'].astype(str).str.contains('unexplained').any(), areas        # Round 1's table was kept.
    assert areas.loc[areas['Group'] != 'Load', 'Balance MW'].abs().sum() > 1, areas

def test_final_balance_failing_inside_its_scaling_keeps_one_row():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.base['Load'].loc[3, 'SMW'] = 900.0                       # Swing at +700 MW.
    real = scale.iterate_to_gen_load_targets
    def boom(*a, **k):
        raise RuntimeError('inside')
    scale.iterate_to_gen_load_targets = boom
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            runs = area.balance_swing(sa)
    finally:
        scale.iterate_to_gen_load_targets = real
    assert [r.name for r in runs] == ['balance1'] and "failed: RuntimeError('inside')" in runs[0].note
    assert runs[0].before is not None and pd.notna(runs[0].after)

def test_step_02_keeps_the_hour_when_the_dist_gen_cleanup_raises():
    runs, areas, out = step_02_hour_with_a_failure(fail_distgen=True)
    row = runs[runs['Run'] == 'distgen'].iloc[0]
    assert row['Note'].startswith("failed: RuntimeError('no dist gen')"), runs
    assert not areas['Note'].astype(str).str.contains('unexplained').any(), areas

# ------------------ T4: non-conforming loads take the EPC's status and MW 1:1 ------------------
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
        mk = []
        area.makeup(sa, targets, tables=mk, nonconforming=reports['nonconforming'])
        sheet = area.areas(area.read_case(sa), reports, targets, [], makeup_tables=mk)
    return sa, targets, reports, sheet, mk

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
    # s1: Open in the seed, Closed in the EPC (40 MW): closed in. s2: Closed in the seed, Open in the EPC: opened.
    # s3: no EPC key: untouched. Dist gen of the 1:1 loads goes Open & 0.
    sa, targets, reports, sheet, _ = nonc_hour(
        [('s1', 1, 'Open', 7.0, 0.0, 0.0, 0.0), ('s2', 1, 'Closed', 60.0, 0.0, 0.0, 0.0), ('s3', 1, 'Closed', 30.0, 0.0, 0.0, 0.0)],
        [('s1', 1, 'Closed', 30.0, 5.0, 5.0, 4.0), ('s2', 1, 'Open', 99.0, 0.0, 0.0, 0.0)],
        flags=('NO', 'NO', 'NO'), epc_flags=('NO', 'NO'))
    t = targets['Load']
    assert t.loc[['s1', 's2'], 'Status_Target'].tolist() == ['Closed', 'Open'] and 's3' not in t.index
    assert t.loc[['s1', 's2'], 'SMW_Target'].tolist() == [40.0, 0.0] and t.loc['s1', 'SMvar_Target'] == 4.0
    assert (t.loc[['s1', 's2'], ['IMW_Target', 'ZMW_Target', 'DistMWInput_Target']] == 0).all().all()
    assert t.loc[['s1', 's2'], 'DistStatus_Target'].tolist() == ['Open', 'Open']
    out = sa.table('Load').set_index('BusNum')
    assert (out.loc[4, 'Status'], out.loc[4, 'SMW'], out.loc[4, 'SMvar'], out.loc[4, 'DistStatus']) == ('Closed', 40.0, 4.0, 'Open')
    assert (out.loc[5, 'Status'], out.loc[5, 'SMW']) == ('Open', 0.0)
    assert (out.loc[6, 'Status'], out.loc[6, 'SMW']) == ('Closed', 30.0)
    row = sheet[sheet['Group'] == 'Load'].iloc[0]
    assert 'non-conforming set to EPC: 40.0 MW (2 loads)' in row['Note'] and 'non-conforming left as is: 30.0 MW' in row['Note']
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
