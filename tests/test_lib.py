"""
Tests for the parts of Scripts/ which don't need PowerWorld. Runs on any OS.
    python tests/test_lib.py      (or: pytest tests)
"""
import re
import sys
from pathlib import Path
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from Scripts import pw, topo, scale

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
        elif command.startswith('SolvePowerFlow') and self.bad & {k for _, k in self.applied}:
            return ('Diverged',)
        return ('',)

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

# ------------------ topo ------------------
def test_transformer_is_okay_flags_low_fixed_taps():
    df = pd.DataFrame({'Tap': [1.0, 1.0], 'NomkVFrom': 230.0, 'XFNomkVbaseFrom': 230.0, 'NomkVTo': 115.0, 'XFNomkVbaseTo': 115.0,
                       'TapFixedFrom': [1.0, 0.5], 'TapFixedTo': 1.0})
    assert topo.transformer_is_okay(df).tolist() == [True, False]

def test_suggested_shunt_status():
    df = pd.DataFrame({'MvarNom': [50, 50, -50, -50, 50], 'Status': ['Open', 'Closed', 'Closed', 'Open', 'Open'], 'Vpu': [0.9, 1.1, 0.9, 1.1, 1.0]})
    assert topo.suggested_shunt_status(df, 0.92, 1.08).tolist() == ['Closed', 'Open', 'Open', 'Closed', 'Open']

def test_branch_status_changes():
    def case(rows):
        df = pd.DataFrame(rows, columns=['ObjectID', 'Status', 'BranchDeviceType'])
        return {'Branch': {'df': df}, 'Transformer': {'df': df.iloc[0:0]}}
    right = case([['a', 'Closed', 'Line'], ['b', 'Closed', 'Line'], ['c', 'Closed', 'Line'], ['f', 'Closed', 'Fuse']])
    left = case([['a', 'Closed', 'Line'], ['b', 'Open', 'Line']])
    _, changes = topo.branch_status_changes(left, right)
    assert changes.set_index('ObjectID')['StatusLeft'].to_dict() == {'b': 'Open', 'c': 'Open'}  # Fuse ignored.

# ------------------ scale ------------------
def test_build_targets():
    common = dict(BusName='A', NomkV=230.0)
    case_gen = pd.DataFrame({'ObjectID': ['g1', 'g2', 'g3'], 'BusNum': [1, 2, 3], 'ID': '1', 'Status': ['Closed', 'Open', 'Closed'], 'MWSetPoint': [10.0, 7.0, 5.0], **common})
    gv_gen = pd.DataFrame({'ObjectID': ['g1', 'g2'], 'BusNum': [1, 2], 'ID': '1', 'Status': ['Closed', 'Open'], 'MWSetPoint': [20.0, 9.0], **common})
    load = lambda ids: pd.DataFrame({'ObjectID': ids, 'BusNum': 1, 'ID': '1', 'Status': 'Closed', 'SMW': 50.0, 'SMvar': 5.0, 'DistStatus': 'Closed', 'DistMWInput': 3.0, 'DistMvarInput': 1.0, **common})
    targets = scale.build_targets({'Gen': gv_gen, 'Load': load(['l1'])}, {'Gen': case_gen, 'Load': load(['l1', 'l2'])})
    gen, ld = targets['Gen'], targets['Load']

    assert gen.loc['g1', 'MWSetPoint_Target'] == 20 and gen.loc['g1', 'MWSetPoint_Seed'] == 10
    assert gen.loc['g2', 'MWSetPoint'] == 0 and gen.loc['g2', 'MWSetPoint_Target'] == 0            # Open now and in target.
    assert gen.loc['g3', 'Status_Target'] == 'Open' and gen.loc['g3', 'MWSetPoint_Target'] == 0    # Missing from target.
    assert ld.loc['l1', 'DistMWInput_Target'] == 0 and ld.loc['l1', 'DistStatus_Target'] == 'Open'  # GridView dist gen unused.
    assert ld.loc['l1', 'DistMWInput'] == 3.0                                                       # Current dist gen kept.
    assert ld.loc['l2', 'Status_Target'] == 'Open' and ld.loc['l2', 'SMW_Target'] == 0
    assert scale.included(gen).all() and scale.included(ld).all()

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

def test_deltas_and_next_step():
    gen = targets_df(ObjectID=['g1', 'g2'], MWSetPoint=[0.0, 0.0], MWSetPoint_Target=[100.0, 100.0])
    gen.loc['g2', 'ExclusionReason'] = 'PVQV'
    scale.compute_deltas({'Gen': gen}, 4)
    step_df = scale.next_step('Gen', gen)
    assert step_df['MWSetPoint'].to_dict() == {'g1': 25.0}
    assert gen['MWSetPoint'].tolist() == [0.0, 0.0]  # Unchanged until applied.

def test_close_all_related_gen_load_skips_excluded_and_restores_failures():
    sa = FakeSimAuto(bad={'2'})  # Closing the gen on bus 2 diverges.
    gen = targets_df(ObjectID=['g1', 'g2', 'g3'], BusNum=[1, 2, 3], ID='1', Status='Open', Status_Target='Closed', MWSetPoint=[0.0, 0.0, 0.0])
    gen.loc['g3', 'ExclusionReason'] = 'PVQV'
    load = targets_df(ObjectID=['l1'], BusNum=[9], ID=['1'], Status=['Open'], Status_Target=['Closed'], DistStatus=['Closed'], DistStatus_Target=['Closed'],
                      SMW=[1.0], SMvar=[0.0], DistMWInput=[0.0], DistMvarInput=[0.0])
    scale.close_all_related_gen_load(sa, {'Gen': gen, 'Load': load})
    assert gen['Status'].tolist() == ['Closed', 'Open', 'Open']
    assert gen['ExclusionReason'].tolist() == ['', 'Diverged closing at 0 MW', 'PVQV']
    assert ('Gen', '3') not in sa.applied
    assert load.loc['l1', 'Status'] == 'Closed' and load.loc['l1', 'SMW'] == 0

def test_iterate_pretest_mode():
    sa = FakeSimAuto(bad={'2'})  # Closing the gen on bus 2 diverges; caught before scaling.
    targets, pvqv = scaling_targets(gen_status=['Closed', 'Open'])
    log = scale.iterate_to_gen_load_targets(sa, targets, pvqv, iterations=4)
    gen, load = targets['Gen'], targets['Load']

    assert log['iteration_df']['Value'].iloc[0] == 'Completed 4 of 4 iterations'
    assert gen.loc['g1', 'MWSetPoint'] == 50.0 and gen.loc['g2', 'Status'] == 'Open'
    assert log['gen']['ExclusionReason'].to_dict() == {'g2': 'Diverged closing at 0 MW'}
    assert load['SMW'].tolist() == [10.0, 8.0] and load.loc['l2', 'Status'] == 'Closed'
    assert sa.states == {}

def test_iterate_stops_at_divergence_without_bisect():
    sa = FakeSimAuto(bad={'2'})  # g2 is already closed, so it first diverges during scaling.
    targets, pvqv = scaling_targets(gen_status='Closed')
    log = scale.iterate_to_gen_load_targets(sa, targets, pvqv, iterations=4)
    assert log['iteration_df']['Value'].iloc[0] == 'Completed 0 of 4 iterations'
    assert targets['Gen']['MWSetPoint'].tolist() == [10.0, 10.0]

def test_iterate_bisect_excludes_divergent_and_continues():
    sa = FakeSimAuto(bad={'2'})
    targets, pvqv = scaling_targets(gen_status='Closed')
    log = scale.iterate_to_gen_load_targets(sa, targets, pvqv, iterations=4, bisect=True)
    gen = targets['Gen']
    assert log['iteration_df']['Value'].iloc[0] == 'Completed 4 of 4 iterations'
    assert gen['MWSetPoint'].tolist() == [50.0, 10.0]  # g2 left at its last good step.
    assert log['gen']['ExclusionReason'].to_dict() == {'g2': 'Diverged during scaling'}
    assert targets['Load']['SMW'].tolist() == [10.0, 8.0]

if __name__ == '__main__':
    tests = [f for name, f in dict(globals()).items() if name.startswith('test_')]
    for test in tests:
        test()
    print(f'{len(tests)} tests passed.')
