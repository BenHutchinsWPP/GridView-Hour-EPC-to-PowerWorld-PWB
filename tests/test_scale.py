"""
Tests for Scripts/scale.py (the scaling engine). Needs no PowerWorld.
    python tests/test_scale.py      (all tests: python tests/run.py [name part])
"""
import contextlib
import io
import tempfile
from pathlib import Path
import numpy as np
import pandas as pd
from helpers import FakeSimAuto, LOAD_P_ALL, SwingSimAuto, gens, loads, results, swing_case, targets_df
from Scripts import pw, scale, area

def scaling_targets(gen_status):
    """g1 on bus 1 and g2 on bus 2 scaling up; l1 scaling up and l2 closing in."""
    gen = targets_df(ObjectID=['g1', 'g2'], BusNum=[1, 2], ID='1', Status=gen_status, Status_Target='Closed',
                     MWSetPoint=[10.0, 10.0], MWSetPoint_Target=[50.0, 20.0])
    load = targets_df(ObjectID=['l1', 'l2'], BusNum=[1, 1], ID=['1', '2'], Status=['Closed', 'Open'], Status_Target='Closed',
                      DistStatus='Open', DistStatus_Target='Open',
                      **{c: [5.0, 0.0] for c in scale.SCALED['Load']}, **{f'{c}_Target': [10.0, 8.0] for c in scale.SCALED['Load']})
    pvqv = pd.DataFrame({'Number': [1, 2], 'Name': 'x', 'NomkV': 230.0, 'SensdVdPself': 0.0, 'SensdVdQself': 0.0})
    return {'Gen': gen, 'Load': load}, pvqv

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

def test_rescale_caps_gen_targets_at_mwmax():
    df = targets_df(ObjectID=['a', 'b', 'c'], MWSetPoint=[50.0, 50.0, 50.0], MWSetPoint_Start=[0.0, 0.0, 0.0], MWSetPoint_Target=[100.0, 100.0, 130.0])
    scale.rescale(df, pd.Series({'a': 1.2, 'b': 1.2, 'c': 1.2}), ['MWSetPoint'], cap=pd.Series({'a': 110.0, 'b': 1000.0, 'c': 120.0}))
    assert df['MWSetPoint_Target'].round(6).tolist() == [110.0, 120.0, 130.0]   # c was already above its MWMax: not raised.
    assert df['MWSetPoint'].round(6).tolist() == [60.0] * 3                     # Current values are not capped.

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

def test_report_gen_load_balance_counts_only_closed_units():
    gen = targets_df(ObjectID=['g1', 'g2'], Status=['Closed', 'Open'], Status_Target=['Open', 'Closed'],
                     MWSetPoint=[100.0, 50.0], MWSetPoint_Target=[100.0, 50.0])
    load = targets_df(ObjectID=['l1', 'l2'], Status=['Closed', 'Open'], Status_Target='Closed', DistStatus='Closed', DistStatus_Target='Open',
                      **{c: [10.0, 10.0] for c in scale.SCALED['Load']}, **{f'{c}_Target': [10.0, 10.0] for c in scale.SCALED['Load']})
    out = scale.report_gen_load_balance({'Gen': gen, 'Load': load})
    assert out['Base Case'].tolist() == [100.0, 10.0, 30.0, 80.0] and out['Target Case'].tolist() == [50.0, 0.0, 60.0, -10.0]

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

def test_seed_islands_are_not_checked_when_the_swing_is_not_in_island_1():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load, slack=[1])
    assert scale.seed_islands(sa, [1]).to_dict() == {1: 1, 2: 1, 3: 1, 99: 1}
    sa.islands = {1: '2'}
    assert scale.seed_islands(sa, [1]) is None
    sa.islands = {2: '2'}                                                       # A live second island with its own slack bus.
    assert scale.seed_islands(sa, [1, 2]).to_dict()[2] == 2

def test_a_unit_on_a_bus_out_of_band_from_the_start_is_held_before_scaling_begins():
    gen = targets_df(ObjectID=['g1', 'g2'], BusNum=[1, 2], ID='1', Status='Closed', Status_Target='Closed',
                     MWSetPoint=[10.0, 10.0], MWSetPoint_Target=[20.0, 20.0])
    gen['HeldReason'] = ''
    scale.hold_voltage({'Gen': gen}, pd.Series({1: 1.0, 2: 0.85}))          # No Progress column yet (before mark_start()).
    assert gen['HeldReason'].tolist() == ['', 'Voltage']

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

if __name__ == '__main__':
    tests = [f for name, f in dict(globals()).items() if name.startswith('test_')]
    for test in tests:
        test()
    print(f'{len(tests)} tests passed.')
