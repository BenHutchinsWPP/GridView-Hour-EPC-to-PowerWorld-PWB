"""
Tests for pw.py, prep.py, topo.py and settings.py. Needs no PowerWorld.
    python tests/test_pw_prep_topo.py      (all tests: python tests/run.py [name part])
"""
import contextlib
import io
import tempfile
from pathlib import Path
import numpy as np
import pandas as pd
from helpers import cell, FakeSimAuto, SwingSimAuto, gens, load_step_02, loads, module_settings, saved_settings, targets_df
from Scripts import pw, prep, scale, area, settings, topo

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

def test_suggested_shunt_status():
    df = pd.DataFrame({'MvarNom': [50, 50, -50, -50, 50], 'Status': ['Open', 'Closed', 'Closed', 'Open', 'Open'], 'Vpu': [0.9, 1.1, 0.9, 1.1, 1.0]})
    assert prep.suggested_shunt_status(df, 0.92, 1.08).tolist() == ['Closed', 'Open', 'Open', 'Closed', 'Open']

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
    saved = (scale.EXCLUDE_VMIN, dict(area.UNIT_GENTYPE), set(area.CAPACITY_TYPES), list(area.GEN_TYPES))
    with tempfile.TemporaryDirectory() as d:
        fp = Path(d) / 's.toml'
        try:
            fp.write_text('[run]\nmode = "topology"\n[scale]\nexclude_vmin = 0.7\n'
                          '[area]\ncapacity_types = ["Wind"]\ngen_types = ["Solar", "Hydro", "Wind", "Thermal", "Battery Storage", "Unknown", "Pumped"]\n'
                          '[area.unit_gentype]\nps = "Pumped"\nHY = "fuel"\n')
            run = settings.load(fp)
            assert run['mode'] == 'topology' and run['balance_swing'] is True     # Default filled in.
            assert scale.EXCLUDE_VMIN == 0.7 and area.CAPACITY_TYPES == {'Wind'}
            assert area.UNIT_GENTYPE['PS'] == 'Pumped' and area.UNIT_GENTYPE['HY'] is None and area.UNIT_GENTYPE['GT'] == 'Thermal'
            # The setting reaches code which reads it at call time.
            gen = targets_df(ObjectID=['g1'], BusNum=[1])
            assert scale.voltage_ok(gen, pd.Series({1: 0.75})).tolist() == [True]
            scale.EXCLUDE_VMIN, area.UNIT_GENTYPE, area.CAPACITY_TYPES, area.GEN_TYPES = saved[0], dict(saved[1]), set(saved[2]), list(saved[3])
            for bad in ['[run]\nmode = "x"\n', '[run]\nmodee = "area"\n', '[scale]\nnope = 1\n', '[nope]\n',
                        '[area.unit_gentype]\nPS = "Pumped"\n', '[run]\nswing_bus = true\n', '[scale]\nexclude_vmin = "low"\n',
                        '[run]\nbalance_swing = 1\n', '[scale]\nstatus = {}\n', '[area]\ncapacity_types = "Wind"\n']:
                fp.write_text(bad)
                try:
                    settings.load(fp)
                    assert False, bad
                except ValueError:
                    pass
        finally:
            scale.EXCLUDE_VMIN, area.UNIT_GENTYPE, area.CAPACITY_TYPES, area.GEN_TYPES = saved

def test_trial_refuses_to_nest():
    sa = FakeSimAuto()
    with pw.trial(sa):
        try:
            with pw.trial(sa):
                pass
            assert False
        except RuntimeError:
            pass

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
        pw.print_opened = pw.print_solve_errors = True       # Importing the step must leave both as they are.
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

def test_giant_swing_on_a_bus_not_in_the_case_is_a_clear_error():
    sa = FakeSimAuto()
    try:
        prep.create_giant_swing(sa, 424242)
        assert False
    except ValueError as e:
        assert '424242' in str(e) and 'not in the case' in str(e)
    assert not sa.applied

def conversion_case(bad=()):
    """Area 1: loads on buses 11 and 12 (ZIP + dist gen), 13 Open with dist gen (S 50, I 7). Area 2: bus 21. Swing xx on bus 99 (slack)."""
    gen = gens([('xx', 9, 'xx', 'Closed', '', 0.0, 2e4, -2e4)]).assign(BusNum=99, MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('a', 1, 'Closed', 60.0, 20.0, 10.0, 6.0), ('b', 1, 'Closed', 40.0, 0.0, 0.0, 4.0),
                  ('c', 1, 'Open', 50.0, 7.0, 0.0, 5.0), ('d', 2, 'Closed', 30.0, 5.0, 5.0, 3.0)]).assign(BusNum=[11, 12, 13, 21], NomkV=230.0, BusName='b')
    load.loc[load['ObjectID'] == 'c', 'DistMWInput'] = 9.0
    sa = SwingSimAuto(gen, load, slack=[99])
    sa.bad, sa.count_dist = set(bad), True
    return sa

def test_constant_power_rows_use_net_load_and_skip_open_loads():
    load = loads([('a', 1, 'Closed', 60.0, 20.0, 10.0, 6.0), ('o', 1, 'Open', 50.0, 0.0, 0.0, 5.0)]).assign(NetMW=[85.0, 0.0], NetMvar=[5.0, 0.0])
    out = prep.constant_power_rows(load)
    assert out.index.tolist() == ['a'] and cell(out, 'a', ['SMW', 'SMvar']).tolist() == [85.0, 5.0]
    assert bool((out.loc['a', ['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']] == 0).all()) and out.loc['a', 'DistStatus'] == 'Open'

def test_convert_loads_to_constant_power_keeps_the_operating_point():
    sa = conversion_case()
    before = sa.table('Load').set_index('ObjectID')
    with contextlib.redirect_stdout(io.StringIO()) as out:
        sheets = prep.convert_loads_to_constant_power(sa)
    after = sa.table('Load').set_index('ObjectID')
    assert set(sheets) == {'load_conversion'} and 'Loads to constant power: 3 loads, net 155.0 MW (dist gen 15.0 MW folded in)' in out.getvalue()
    assert 'WARNING' not in out.getvalue()
    for k in 'abd':          # Written S is the net load before; the rest is 0, and dist gen is Open (not just 0).
        assert after.loc[k, 'SMW'] == before.loc[k, 'NetMW'] and after.loc[k, 'SMvar'] == before.loc[k, 'NetMvar']
        assert (after.loc[k, ['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']] == 0).all() and after.loc[k, 'DistStatus'] == 'Open'
    assert after.loc['a', 'SMW'] == 85.0
    # The Open load keeps S, I and Z, and loses only its dist gen.
    assert after.loc['c', ['SMW', 'IMW', 'ZMW', 'SMvar']].tolist() == before.loc['c', ['SMW', 'IMW', 'ZMW', 'SMvar']].tolist()
    assert after.loc['c', 'DistStatus'] == 'Open' and after.loc['c', 'DistMWInput'] == 0 and after.loc['c', 'DistMvarInput'] == 0
    sheet = sheets['load_conversion'].set_index('AreaNumber')
    assert sheet.loc[1, 'Loads'] == 2 and sheet.loc[1, 'Nominal MW'] == 130.0 and sheet.loc[1, 'Dist MW'] == 10.0
    assert (sheet[['Diff MW', 'Diff Mvar']].dropna().abs() < 1e-6).all().all()
    assert abs(cell(sheet, 'Swing', 'Net MW Before') - cell(sheet, 'Swing', 'Net MW After')) < 1e-6 and sheet.loc['Swing', 'Net MW Before'] == 155.0

def test_convert_loads_lists_loads_which_diverge_and_converts_the_rest():
    sa = conversion_case(bad={'12'})      # Bus 12 (load b) diverges once changed.
    with contextlib.redirect_stdout(io.StringIO()) as out:
        sheets = prep.convert_loads_to_constant_power(sa)
    after = sa.table('Load').set_index('ObjectID')
    assert sheets['load_conversion_failed'].index.tolist() == ['b'] and 'WARNING: 1 loads could not be converted' in out.getvalue()
    assert after.loc['b', 'DistStatus'] == 'Closed' and after.loc['b', 'DistMWInput'] == 5.0         # Untouched.
    assert after.loc['a', 'SMW'] == 85.0 and after.loc['a', 'IMW'] == 0 and after.loc['d', 'IMW'] == 0

def test_convert_loads_stops_bisecting_at_the_failure_cap():
    sa = conversion_case(bad={'11', '12'})
    saved = prep.CONVERT_MAX_FAILED
    prep.CONVERT_MAX_FAILED = 1
    try:
        with contextlib.redirect_stdout(io.StringIO()) as out:
            sheets = prep.convert_loads_to_constant_power(sa)
    finally:
        prep.CONVERT_MAX_FAILED = saved
    after = sa.table('Load').set_index('ObjectID')
    assert 'stopped bisecting after 1 failures' in out.getvalue()
    assert 'd' in sheets['load_conversion_failed'].index and after.loc['d', 'IMW'] == 5.0      # d is fine, but was never tried.
    # Exactly one single-load conversion was tried: the cap stops at the first failure, not after it.
    sa, calls, real = conversion_case(bad={'11', '12'}), [], pw.set_df
    def counting(S, t, df, f, *a, **k):
        if t == 'Load' and 'SMW' in f:
            calls.append(len(df))
        return real(S, t, df, f, *a, **k)
    prep.CONVERT_MAX_FAILED, pw.set_df = 1, counting
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            prep.convert_loads_to_constant_power(sa)
    finally:
        prep.CONVERT_MAX_FAILED, pw.set_df = saved, real
    assert calls.count(1) == 1, calls

def test_conversion_warnings_flag_area_swing_and_load_moves():
    sheet = pd.DataFrame({'AreaNumber': [1, 2, 'Swing'], 'Diff MW': [0.5, 0.0, 0.2], 'Diff Mvar': [0.0, 3.0, np.nan]})
    done = pd.DataFrame({'MW': [10.0, 20.0], 'SMW': [10.0, 20.0], 'DistMW': [0.0, 0.0]})
    w = prep.conversion_warnings(sheet, done)
    assert len(w) == 1 and 'area(s) 2.' in w[0]
    sheet.loc[2, 'Diff MW'] = -5.0
    done.loc[1, 'MW'] = 17.0
    w = prep.conversion_warnings(sheet.assign(**{'Diff Mvar': [0.0, 0.0, np.nan]}), done)
    assert len(w) == 2 and 'swing moved -5.0 MW' in w[0] and '1 converted loads' in w[1]
    assert prep.conversion_warnings(sheet.assign(**{'Diff MW': 0.0, 'Diff Mvar': 0.0}), done.assign(MW=done['SMW'])) == []

def test_convert_loads_with_blank_net_fields_stops_before_writing():
    sa = conversion_case()
    real = sa.GetParametersMultipleElementRect
    setattr(sa, 'GetParametersMultipleElementRect',
            lambda t, f, g: ('', [['' if x == 'NetMW' else v for x, v in zip(f, row)] for row in real(t, f, g)[1]]) if t == 'Load' else real(t, f, g))
    try:
        prep.convert_loads_to_constant_power(sa)
        assert False, 'expected SystemExit'
    except SystemExit as e:
        assert '3 Closed loads have a blank NetMW or NetMvar' in str(e)
    assert not any(t == 'Load' for t, _ in sa.applied)

def test_convert_loads_skips_the_conversion_when_the_case_does_not_solve():
    sa = conversion_case(bad={'11'})
    sa.applied[('Load', '11')] = {}       # Already diverged before the conversion.
    with contextlib.redirect_stdout(io.StringIO()) as out:
        sheets = prep.convert_loads_to_constant_power(sa)
    assert sheets == {} and 'NOT converted' in out.getvalue() and set(sa.applied) == {('Load', '11')}

def test_step_01_runs_the_conversion_after_agc_off_and_before_pvqv_and_the_giant_swing():
    lines = (Path(__file__).resolve().parents[1] / 'pcm_to_pf.py').read_text().splitlines()
    first = lambda text: next(i for i, line in enumerate(lines) if text in line)
    order = [first(f'prep.{name}') for name in ('disable_area_agc', 'convert_loads_to_constant_power', 'get_pvqv', 'create_giant_swing')]
    assert order == sorted(order)

class ThingSimAuto(FakeSimAuto):
    """A 'Thing' table of the given rows [ObjectID, R, I, S]; counts GetFieldList calls and records the rows written."""
    def __init__(self, rows, field_rows=None):
        super().__init__()
        self.rows, self.field_calls, self.written = rows, 0, []
        self.field_rows = field_rows if field_rows is not None else [
            ['Thing:R', 'R', 'Real', 'r', 'R', 'YES'], ['Thing:I', 'I', 'Integer', 'i', 'I', 'YES'], ['Thing:S', 'S', 'String', 's', 'S', 'YES']]

    def GetFieldList(self, table):
        self.field_calls += 1
        return ('', self.field_rows)

    def GetParametersMultipleElementRect(self, table, fields, filter_group):
        names = ['ObjectID', 'R', 'I', 'S']
        return ('', [[row[names.index(f)] for f in fields] for row in self.rows])

    def ChangeParametersMultipleElementRect(self, table, fields, rows):
        self.written = rows
        return ('',)

def test_blank_values_round_trip_through_get_df_and_set_df():
    pw.clear_field_types()
    rows = [['a', '1.5', '101', 'x'], ['b', '', '', ''], ['c', '2.5', '7', 'y']]
    sa = ThingSimAuto(rows)
    df = pw.get_df(sa, 'Thing', ['ObjectID', 'R', 'I', 'S'])
    assert all(type(v) is str for v in [*df['S'], *df['ObjectID']])
    assert str(df['R'].dtype) == 'float64' and str(df['I'].dtype) == 'Int64'
    assert np.isnan(df.loc[1, 'R']) and df.loc[1, 'I'] is pd.NA and df.loc[1, 'S'] == ''
    pw.set_df(sa, 'Thing', df, ['ObjectID', 'R', 'I', 'S'])
    assert sa.written == rows
    pw.clear_field_types()

def test_field_types_keys_by_variable_name_and_concise_name_and_falls_back_to_the_base_name():
    pw.clear_field_types()
    sa = ThingSimAuto([['a', '1', '2', 'x']], field_rows=[
        ['Thing:R', 'RealVarName', 'Real', 'd', 'RConcise', 'YES'], ['Thing:I', 'CustomInteger', 'Integer', 'd', '', 'YES'],
        ['Thing:S', 'Odd', 'Weird', 'd', 'SConcise', 'YES'], ['Thing:N', 'NoConcise', 'Real']])
    assert pw.field_types(sa, 'Thing') == {'realvarname': 'Real', 'rconcise': 'Real', 'custominteger': 'Integer',
                                           'odd': 'String', 'sconcise': 'String', 'noconcise': 'Real'}
    assert pw._field_type(sa, 'Thing', 'CustomInteger:0') == 'Integer' and pw._field_type(sa, 'Thing', 'RConcise') == 'Real'
    assert pw._field_type(sa, 'Thing', 'realVARname') == 'Real'      # PowerWorld's names ignore case.
    pw.clear_field_types()
    # A concise name ([4]) never overrides another field's variable name ([1]).
    sa = ThingSimAuto([], field_rows=[['Thing:A', 'A', 'Real', 'd', 'B', 'YES'], ['Thing:B', 'B', 'String', 'd', '', 'YES']])
    assert pw.field_types(sa, 'Thing') == {'a': 'Real', 'b': 'String'}
    assert pw._field_type(sa, 'Thing', 'ObjectID') == 'String'
    pw.clear_field_types()

def test_a_field_missing_from_the_field_list_stops_the_read():
    pw.clear_field_types()
    sa = ThingSimAuto([['a', '1.5', '101', 'x']])
    try:
        pw.get_df(sa, 'Thing', ['ObjectID', 'R', 'Mystery', 'Other:1'])
        assert False
    except pw.SimAutoError as e:
        assert 'GetFieldList(Thing) has no type for Mystery, Other:1.' in str(e)
    # An Integer field with a fraction: a SimAutoError, not a TypeError.
    try:
        pw.get_df(ThingSimAuto([['a', '1', '1.5', 'x']]), 'Thing', ['I'])
        assert False
    except pw.SimAutoError as e:
        assert 'Thing.I' in str(e)
    pw.clear_field_types()

def test_get_field_list_is_called_once_per_table():
    pw.clear_field_types()
    sa = ThingSimAuto([['a', '1.5', '101', 'x']])
    pw.get_df(sa, 'Thing', ['ObjectID', 'R'])
    pw.get_df(sa, 'Thing', ['I', 'S'])
    assert sa.field_calls == 1
    pw.clear_field_types()

def test_settings_check_steps_and_workers():
    with saved_settings(), tempfile.TemporaryDirectory() as d:
        fp = Path(d) / 's.toml'
        fp.write_text('[run]\nsteps = ["scale"]\nworkers = 4\n')
        run = settings.load(fp)
        assert run['steps'] == ['scale'] and run['workers'] == 4
        for text in ('[run]\nsteps = ["scal"]\n', '[run]\nsteps = []\n', '[run]\nworkers = 0\n', '[run]\nworkers = 1.5\n'):
            fp.write_text(text)
            try:
                settings.load(fp)
                assert False, text
            except ValueError:
                pass

if __name__ == '__main__':
    tests = [f for name, f in dict(globals()).items() if name.startswith('test_')]
    for test in tests:
        test()
    print(f'{len(tests)} tests passed.')
