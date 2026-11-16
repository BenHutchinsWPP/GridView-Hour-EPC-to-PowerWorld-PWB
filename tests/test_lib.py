"""
Tests for the parts of Scripts/ which don't need PowerWorld. Runs on any OS.
    python tests/test_lib.py      (or: pytest tests)
"""
import importlib.util
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

def test_adjust_shunts_writes_nothing_when_all_in_band():
    class ShuntSimAuto(FakeSimAuto):
        def GetParametersMultipleElementRect(self, table, fields, filter_group):
            if table != 'Shunt':
                return super().GetParametersMultipleElementRect(table, fields, filter_group)
            row = {'ObjectID': 'Shunt 1 1', 'MvarNom': '50', 'Status': 'Closed', 'Vpu': '1.0'}
            return ('', [[row[f] for f in fields]])
    sa = ShuntSimAuto()
    prep.adjust_shunts(sa)
    assert sa.applied == {}     # Was: one blank row per shunt, written by a blank ObjectID.

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

def test_open_load_cleanup():
    case = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 10.0), ('lo', 1, 'Open', 50.0, 0.0, 0.0, 0.0)])
    cleanup = area.open_load_cleanup(case)
    assert cleanup['ObjectID'].tolist() == ['lo'] and cleanup['DistMWInput'].tolist() == [0.0]
    assert len(area.open_load_cleanup(case.assign(AGC='NO'))) == 0     # Non-conforming: untouched.

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

def test_balance_targets_half_gen_half_load():
    # Swing at +100 MW: gens up 50 MW, load down 50 MW, each by one factor over the whole case.
    gen = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 300.0, 0.0), ('g2', 2, '2', 'Closed', 'HY', 100.0, 120.0, 0.0),
                ('neg', 2, '3', 'Closed', 'BA', -10.0, 50.0, -50.0), ('sw', 3, '1', 'Closed', 'GT', 500.0, 1e4, 0.0),
                ('dc', 3, '2', 'Closed', 'DC', 90.0, 500.0, -500.0)]).assign(BusNum=[1, 2, 2, 9, 3])
    load = loads([('l1', 1, 'Closed', 150.0, 0.0, 0.0, 30.0), ('l2', 2, 'Closed', 350.0, 0.0, 0.0, 70.0),
                  ('lneg', 2, 'Closed', -40.0, 0.0, 0.0, 0.0), ('lo', 2, 'Open', 99.0, 0.0, 0.0, 0.0)])
    t = area.balance_targets({'Gen': gen, 'Load': load}, swing=100.0, swing_buses=[9])
    assert t['Gen']['MWSetPoint_Target'].to_dict() == {'g1': 130.0, 'g2': 120.0}  # g2 capped at MWMax; g1 takes the rest.
    assert t['Load']['SMW_Target'].to_dict() == {'l1': 135.0, 'l2': 315.0}
    assert t['Load']['SMvar_Target'].to_dict() == {'l1': 27.0, 'l2': 63.0}
    assert t['Load']['P_Seed'].to_dict() == {'l1': 150.0, 'l2': 350.0}     # What rescale_targets() divides by.
    assert (t['Load']['Status_Target'] == 'Closed').all()

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
    assert t['Gen']['MWSetPoint_Target'].to_dict() == {'g1': 150.0}     # The whole gen half lands on what is left.
    assert t['Load']['SMW_Target'].to_dict() == {'l1': 50.0}

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

def test_nonconforming_loads_are_left_alone():
    # l2 / e2 are non-conforming (AGC NO): kept at seed value, out of both area totals, out of balancing.
    case = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 10.0), ('l2', 1, 'Closed', 500.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO '])
    epc = loads([('e1', 1, 'Closed', 150.0, 0.0, 0.0, 15.0), ('e2', 1, 'Closed', 900.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO'])
    targets, reports = area.build_targets({'Load': epc, 'Gen': gens([])}, {'Load': case, 'Gen': gens([])})
    assert targets['Load'].index.tolist() == ['l1'] and targets['Load'].loc['l1', 'SMW_Target'] == 150
    r = reports['area_load'].iloc[0]
    assert (r['Target MW'], r['Seed Non-conforming MW'], r['EPC Non-conforming MW']) == (150, 500, 900)
    bal = area.balance_targets({'Gen': gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 300.0, 0.0)]), 'Load': case}, 50.0, [])
    assert bal['Load'].index.tolist() == ['l1']
    area.SCALE_NONCONFORMING = True
    try:
        targets, _ = area.build_targets({'Load': epc, 'Gen': gens([])}, {'Load': case, 'Gen': gens([])})
        assert targets['Load']['SMW_Target'].sum() == 1050
    finally:
        area.SCALE_NONCONFORMING = False

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
            df['MW'] = df['MWSetPoint']
            sw = (df['ID'] == 'xx') | df['BusNum'].isin(self.slack)
            df.loc[sw, 'MW'] = p - df.loc[on & ~sw, 'MWSetPoint'].sum()
        return df

    def GetParametersMultipleElementRect(self, table, fields, filter_group):
        if table in self.base:
            return ('', [[str(v) for v in row] for row in self.table(table)[fields].itertuples(index=False)])
        if table == 'Bus':
            buses = sorted(set(self.base['Gen']['BusNum']) | set(self.base['Load']['BusNum']))
            values = {'Vpu': '1.0', 'IslandNumber': '1', 'BusIsStarBus:1': 'NO', 'BusNomVolt': '230'}
            return ('', [[str(n) if f in ('Busnum', 'Number') else 'YES' if f == 'Slack' and n in self.slack else values.get(f, 'NO' if f == 'Slack' else '0')
                          for f in fields] for n in buses])
        return ('', [])

def swing_case():
    gen = gens([('xx', 9, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0),
                ('g2', 1, '1', 'Closed', 'HY', 100.0, 1000.0, 0.0)]).assign(BusNum=[99, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
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
    # The load's target was scaled down by the balances, so the scaling ended where the balance left it.
    assert targets['Load'].loc['l1', 'SMW_Target'] < 3200 and abs(targets['Load'].loc['l1', 'SMW'] - targets['Load'].loc['l1', 'SMW_Target']) < 1e-6
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
    moved = {'Gen': targets_df(ObjectID=['g1', 'gx'], MWSetPoint=[60.0, 5.0], MWSetPoint_Seed=[50.0, 4.0], MWMax=[500.0, 500.0]),
             'Load': targets_df(ObjectID=['l1'], SMW=[24.0], IMW=[0.0], ZMW=[0.0], P_Seed=[30.0])}
    area.rescale_targets(targets, moved)
    assert targets['Gen'].loc['g1', ['MWSetPoint', 'MWSetPoint_Start', 'MWSetPoint_Target']].tolist() == [60.0, 0.0, 120.0]
    assert targets['Gen'].loc['g2', 'MWSetPoint_Target'] == 10.0                    # Not moved.
    assert targets['Load'].loc['l1', ['SMvar', 'SMvar_Start', 'SMvar_Target']].tolist() == [8.0, 8.0, 16.0]

# ------------------ review fixes ------------------
@contextmanager
def saved_settings():
    """Module settings (upper-case globals) as they are now, put back afterwards."""
    saved = {(m, k): v for m in settings.MODULES.values() for k, v in vars(m).items() if k.isupper()}
    try:
        yield
    finally:
        for (m, k), v in saved.items():
            setattr(m, k, v)

def test_nonconforming_is_decided_by_key():
    # l2 is non-conforming in the PWB only. The EPC's l2 (500 MW) must not be loaded onto l1.
    case = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 0.0), ('l2', 1, 'Closed', 500.0, 0.0, 0.0, 0.0)]).assign(AGC=['YES', 'NO'])
    epc = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 0.0), ('l2', 1, 'Closed', 500.0, 0.0, 0.0, 0.0)])
    targets, reports = area.build_targets({'Load': epc, 'Gen': gens([])}, {'Load': case, 'Gen': gens([])})
    assert targets['Load'].index.tolist() == ['l1'] and targets['Load'].loc['l1', 'SMW_Target'] == 100
    r = reports['area_load'].iloc[0]
    assert (r['Target MW'], r['Planned MW'], r['Seed Non-conforming MW'], r['EPC Non-conforming MW']) == (100, 100, 500, 500)

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
    assert abs(first['Load MW Moved'] + 50) < 1e-6 and abs(first['Gen MW Moved'] - 50) < 1e-6   # Half each, I part included.

def test_iterate_raises_when_the_case_does_not_solve():
    class Diverges(FakeSimAuto):
        def diverged(self):
            return True
    targets, _ = scaling_targets(gen_status='Closed')
    for sa in (Diverges(),):
        try:
            scale.iterate_to_gen_load_targets(sa, targets)
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
    spec.loader.exec_module(step)
    step.run['mode'] = mode
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
        assert sheets['runs']['Run'].iloc[0] == 'main'
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

def test_slack_bus_gens_are_the_swing_when_there_is_no_giant_swing():
    # Bus 9 is the slack bus: its unit carries the imbalance, and neither balance writes to it.
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0), ('g1', 1, '2', 'Closed', 'GT', 100.0, 1000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 1000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    sa = SwingSimAuto(gen, load, slack=[9])
    assert prep.slack_buses(sa) == [9]
    after_step, log = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
    assert len(log) >= 1 and ('Gen', '9') not in sa.applied
    area.balance_swing(sa)
    assert abs(area.swing_mw(sa.table('Gen'), [9])[0]) < area.BALANCE_TOL_MW and ('Gen', '9') not in sa.applied

def ledger_run(mutate=None, slack=(99,)):
    """Midrun + final balances, a diverging gen (g3, bus 4) and the areas sheet of the result. slack: slack buses (99 has xx)."""
    gen = gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0),
                ('g2', 1, '2', 'Closed', 'HY', 100.0, 1000.0, 0.0), ('g3', 1, '3', 'Closed', 'GT', 100.0, 1000.0, 0.0)]
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
    tables = []
    area.balance_swing(sa, targets, tables=tables)
    assert midrun_log and tables
    return area.areas(area.read_case(sa), reports, targets, tables, swing_buses=list(slack)), targets

def test_areas_ledger_holds_with_midrun_and_final_balances_and_a_diverging_gen():
    out, targets = ledger_run()
    assert out.columns.tolist() == area.AREAS_COLUMNS and out.groupby(['AreaNumber', 'Group']).size().max() == 1
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Balance MW']
    assert (abs(out['Final MW'] - ledger) < 1e-6).all(), out
    row = lambda g: out[out['Group'] == g].iloc[0]
    assert row('Thermal')['Not Reached MW'] > 0 and row('Thermal')['Balance MW'] != 0     # g3 held at 100 of 200 MW; g1 balanced.
    assert row('Load')['Balance MW'] != 0 and pd.isna(row('Thermal')['Target Mvar']) and row('Load')['Target Mvar'] == 320
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
    saved = {(m, k): v for m in settings.MODULES.values() for k, v in vars(m).items() if k.isupper()}
    try:
        run = settings.load()
        assert run['mode'] in settings.MODES
    finally:
        for (m, k), v in saved.items():
            setattr(m, k, v)

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

def test_step_02_saves_the_case_when_the_log_cannot_be_written():
    gen, load = two_cases()
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (gen, load), 'Prepared.pwb': (gen, load)}, slack=[99])
        real = pw.write_excel
        def locked(fp, sheets):
            if fp.name.endswith('_Log.xlsx') and fp.exists():     # Open in Excel at the final write.
                raise PermissionError(fp)
            return real(fp, sheets)
        step.pw.write_excel = locked
        try:
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
        finally:
            step.pw.write_excel = real
        assert 'h1.pwb' in sa.saved

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
    assert ex[['Table', 'ObjectID', 'BusNum', 'Reason']].values.tolist() == [['STATCOM', '7 SC', 7, 'midrun1: STATCOM added']]

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
        runs_b = pd.DataFrame({'Run': ['main'], 'Swing MW After': [7.0]}).reindex(columns=scale.RUN_COLUMNS)
        pw.write_excel(out_dir / 'HourB_Log.xlsx', {'areas': pd.DataFrame(columns=area.AREAS_COLUMNS),
                                                    'exceptions': pd.DataFrame(), 'runs': runs_b})
        target = merge_reports.merge(out_dir)
        sheets = pd.read_excel(target, sheet_name=None)
    assert set(sheets) == {'hours', 'areas', 'exceptions', 'runs'}, set(sheets)
    for name in ('areas', 'exceptions', 'runs'):
        assert sheets[name].columns[0] == 'Hour', name
    assert len(sheets['hours']) == 2
    hours = sheets['hours'].set_index('Hour')
    assert hours.loc['HourA', 'Final Swing MW'] == 25.0 and hours.loc['HourB', 'Final Swing MW'] == 7.0
    assert (hours.loc['HourA', 'Balance Gen MW'], hours.loc['HourA', 'Balance Load MW']) == (3.0, 2.0)
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
    assert not any('Attempting to set' in l or 'Could not set this' in l or 'Opened:' in l for l in lines), lines
    assert any('Main scaling: 0 -> ' in l and 'of 100' in l for l in lines), lines

def test_areas_ledger_holds_with_a_second_slack_bus():
    # Bus 2 (hydro g2) is a slack bus beside the giant swing: left out of the areas, its balance moves must be too.
    out, _ = ledger_run(slack=(99, 2))
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Balance MW']
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

def test_step_02_keeps_the_saved_case_when_its_log_cannot_be_built():
    gen, load = two_cases()
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (gen, load), 'Prepared.pwb': (gen, load)}, slack=[99])
        real = step.area.areas
        def broken(final, *args, **kwargs):
            if final is not None:
                raise ValueError('bug in the log')
            return real(final, *args, **kwargs)
        step.area.areas = broken
        try:
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')   # Must not raise.
        finally:
            step.area.areas = real
    assert 'h1.pwb' in sa.saved

if __name__ == '__main__':
    tests = [f for name, f in dict(globals()).items() if name.startswith('test_')]
    for test in tests:
        test()
    print(f'{len(tests)} tests passed.')
