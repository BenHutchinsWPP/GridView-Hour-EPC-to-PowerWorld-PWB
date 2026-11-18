"""
Tests for pcm_to_pf.py: create_case, merge, main, workers, progress and Prepared_Settings. Needs no PowerWorld.
    python tests/test_steps.py      (all tests: python tests/run.py [name part])
"""
import contextlib
import importlib.util
import io
import json
import tempfile
from pathlib import Path
import numpy as np
import pandas as pd
from helpers import cell, FakeSimAuto, SwingSimAuto, gens, load_step_02, loads, module_settings, restore_settings, saved_settings
from Scripts import pw, prep, scale, area

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
        assert set(sheets) == {'areas', 'exceptions', 'runs', 'shortfall'}
        assert sheets['shortfall'].empty and 'PWB ID' in sheets['shortfall']     # Every group's PWB units cover it.
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

def test_merge_reports_merges_hour_logs():
    spec = importlib.util.spec_from_file_location('merge_reports', Path(__file__).resolve().parents[1] / 'pcm_to_pf.py')
    assert spec is not None and spec.loader is not None
    merge_reports = importlib.util.module_from_spec(spec)
    saved = module_settings()
    spec.loader.exec_module(merge_reports)      # Loads settings.toml: keep the owner's choices out of the tests.
    restore_settings(saved)
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
    assert set(sheets) == {'hours', 'areas', 'exceptions', 'runs', 'shortfall'}, set(sheets)
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
    assert bool(sheets['areas'].set_index('Group').loc['Thermal', ['Seed Loss MW', 'Est. Loss MW', 'Final Loss MW']].isna().all())

def test_step_02_warns_once_when_prepared_pwb_is_unconverted():
    gen, load = two_cases()                                   # loads() fixtures carry dist gen on Closed loads.
    epc_gen = gens([('e1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0)]).assign(BusNum=1, MW=0.0, NomkV=230.0, BusName='b')
    epc_load = loads([('e1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        sa = CaseSimAuto({'h1.epc': (epc_gen, epc_load), 'Prepared.pwb': (gen, load)}, slack=[99])
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            step.create_case(sa, step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb')
    assert sum('Rerun step 01' in l for l in out.getvalue().splitlines()) == 1

def test_step_02_excludes_units_outside_the_main_island_before_the_prebalance():
    # h1 (bus 2) is in island 2: excluded with that reason, turned off before scaling, not counted by the pre-balance.
    # The EPC's e2 (50 MW Hydro) sits on bus 2 too: left out of the targets, so Hydro has Target 0, Unmet 0, and the Note says both.
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
    assert (hydro['Seed MW'], hydro['Target MW'], hydro['Unmet MW'], hydro['Final MW']) == (0.0, 0.0, 0.0, 0.0)
    assert 'outside main island, turned off: 100.0 MW' in hydro['Note'] and 'unexplained' not in hydro['Note']
    assert 'EPC outside main island, left out: 50.0 MW' in hydro['Note']

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
    assert 'gen asked +129 MW, placed +129 MW' in cell(runs, 'makeup', 'Note')     # No other Thermal unit: h1 (Hydro, same area) takes it.
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
    assert 'h1.pwb' in sa.saved and cell(runs, 'makeup', 'Note').startswith('failed: ZeroDivisionError')
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'topology')
        assert step.area.MAKEUP                                                  # On, but area mode only.

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
    runs, areas, _ = step_02_hour_with_a_failure(fail_balance=True)
    assert runs['Run'].tolist()[-2:] == ['balance1', 'balance2'], runs                 # Round 1 is kept...
    assert pd.to_numeric(runs['Swing MW After'], errors='coerce').notna().iloc[-2], runs  # ...with its After.
    assert runs['Note'].iloc[-1].startswith("failed: RuntimeError('boom')"), runs
    assert not areas['Note'].astype(str).str.contains('unexplained').any(), areas        # Round 1's table was kept.
    assert areas.loc[areas['Group'] != 'Load', 'Balance MW'].abs().sum() > 1, areas

def test_step_02_keeps_the_hour_when_the_dist_gen_cleanup_raises():
    runs, areas, _ = step_02_hour_with_a_failure(fail_distgen=True)
    row = runs[runs['Run'] == 'distgen'].iloc[0]
    assert row['Note'].startswith("failed: RuntimeError('no dist gen')"), runs
    assert not areas['Note'].astype(str).str.contains('unexplained').any(), areas

def test_main_runs_the_steps_asked_for_and_skips_a_prepared_seed():
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        calls = []
        def prepare():
            calls.append('prepare')
            step.pw_fp.write_text('')
            step.prep_stamp_fp.write_text(json.dumps(step.prep_stamp()))
        step.prepare = prepare
        step.scale_hours = lambda workers: calls.append(('scale', workers))
        step.merge = lambda out_dir: calls.append('merge')
        step.pw_fp.unlink()
        with contextlib.redirect_stdout(io.StringIO()):
            step.main([])                                   # [run] steps and workers.
            assert calls == ['prepare', ('scale', 1), 'merge']
            calls.clear()                                   # Prepared from these settings: step 01 skipped.
            step.run['workers'] = 3
            step.main(['merge', 'scale'])
            assert calls == [('scale', 3), 'merge']
            calls.clear()
            step.main(['prepare'])
            assert calls == []
            for bad in (['bogus'], ['--workers', '3']):        # Workers: settings.toml only.
                try:
                    with contextlib.redirect_stderr(io.StringIO()):
                        step.main(bad)
                    assert False, bad
                except SystemExit:
                    pass

def test_scale_hours_skips_done_hours_and_hands_the_rest_to_the_workers():
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        for h in ('h2', 'h3'):
            (step.gv_dir / f'{h}.epc').write_text('')
        (step.out_dir / 'h2.pwb').write_text('')           # Done already.
        handed = []
        step._scale_hours_parallel = lambda todo, workers: handed.append(([g.name for g, _ in todo], workers))
        with contextlib.redirect_stdout(io.StringIO()) as out:
            step.scale_hours(8)
        assert handed == [(['h1.epc', 'h3.epc'], 2)] and 'Skipping h2.epc' in out.getvalue()   # No more workers than hours.

def test_a_worker_writes_the_hour_console_and_reports_a_failure():
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        console = step.out_dir / 'h1_Console.txt'
        def fails(SimAuto, gv_fp, out_fp):
            print('scaling h1')
            raise ValueError('diverged')
        step.create_case = fails
        real_connect = pw.connect
        pw.connect = FakeSimAuto
        try:
            assert step._scale_one(step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb', console) == (None, "ValueError('diverged')")
            text = console.read_text()
            assert 'scaling h1' in text and 'Traceback' in text
            step.create_case = lambda SimAuto, gv_fp, out_fp: 'h1: swing 0.0 MW'
            assert step._scale_one(step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb', console) == ('h1: swing 0.0 MW', None)
            def no_seat():
                raise pw.SimAutoError('no free SimAuto seat')
            pw.connect = no_seat                                # A failed connect costs this hour only.
            assert step._scale_one(step.gv_dir / 'h1.epc', step.out_dir / 'h1.pwb', console) == (None, "SimAutoError('no free SimAuto seat')")
        finally:
            pw.connect = real_connect

def test_step_01_is_redone_when_prepared_pwb_was_made_from_other_settings():
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        step.seed_dir.mkdir()
        (step.seed_dir / 'seed.pwb').write_text('seed')
        runs = []
        def prepare():
            runs.append(step.run['mode'])
            step.pw_fp.write_text('prepared')
            step.prep_stamp_fp.write_text(json.dumps(step.prep_stamp()))
        step.prepare = prepare
        with contextlib.redirect_stdout(io.StringIO()) as out:
            assert step.prepared_changes() == ['no Prepared_Settings.json records what it was made from']
            step.check_inputs()                             # An old Prepared.pwb without a record: a warning only.
            step.prepare_if_needed()                        # ...but step 01 redoes it.
            assert runs == ['area'] and step.prepared_changes() == []
            step.prepare_if_needed()                        # The same settings: skipped.
            assert runs == ['area']
            step.check_inputs()
            (step.out_dir / 'h1.pwb').write_text('')
            step.run['mode'] = 'topology'                  # Switching modes is one line in settings.toml.
            changes = step.prepared_changes()
            assert any(c.startswith("mode: 'area' -> 'topology'") for c in changes) and any(c.startswith('epcs') for c in changes)
            try:
                step.check_inputs()                         # Scale alone on a stale Prepared.pwb: stops.
                assert False
            except SystemExit as e:
                assert 'other settings' in str(e)
            step.prepare_if_needed()
            assert runs == ['area', 'topology'] and step.prepared_changes() == []
            assert '1 hours in' in out.getvalue()           # The hours scaled from the old one are pointed out.
            prep.SHUNT_VLOW = 0.9                           # A [prep] setting...
            assert any(c.startswith('prep:') for c in step.prepared_changes())
            prep.SHUNT_VLOW = 0.92
            (step.seed_dir / 'seed.pwb').write_text('a new seed')    # ...or the seed changed.
            assert any(c.startswith('seed:') for c in step.prepared_changes())
            def fails():
                raise SystemExit('the seed does not solve')
            step.prepare = fails
            try:
                step.prepare_if_needed()
            except SystemExit:
                pass
            assert not step.pw_fp.exists() and not step.prep_stamp_fp.exists()     # Nothing stale left behind.

def test_the_progress_bar_reads_each_hours_stage_from_its_console():
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        step = load_step_02(Path(d), 'area')
        fp = Path(d) / 'h1_Console.txt'
        assert step.hour_progress(fp) == (0, 'starting')                     # No console yet.
        lines = ['=== h1.epc ===', 'Seed gen 200 / load 200 MW; target gen 200 / load 300 MW']
        fp.write_text('\n'.join(lines))
        assert step.hour_progress(fp) == (0, 'planning')
        lines += ['Main scaling: 0 -> 25 of 100: solved', 'Main scaling: 25 -> 75 of 100: diverged, halving']
        fp.write_text('\n'.join(lines))
        assert step.hour_progress(fp) == (step.MAIN_SHARE * 75 // 100, 'Main scaling')
        lines += ['Balance 1 (main scaling at 50 of 100): 0 -> 100 of 100: solved']        # A midrun balance.
        fp.write_text('\n'.join(lines))
        assert step.hour_progress(fp) == (step.MAIN_SHARE * 50 // 100, 'Balance 1')
        lines += ['Main scaling: 50 -> 100 of 100: solved', 'Main scaling: done', 'Final balance 2: 0 -> 100 of 100: solved']
        fp.write_text('\n'.join(lines))
        assert step.hour_progress(fp) == (step.MAIN_SHARE, 'Final balance 2')

if __name__ == '__main__':
    tests = [f for name, f in dict(globals()).items() if name.startswith('test_')]
    for test in tests:
        test()
    print(f'{len(tests)} tests passed.')
