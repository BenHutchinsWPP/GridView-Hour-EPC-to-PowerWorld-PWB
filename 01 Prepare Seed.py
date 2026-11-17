"""
Step 01: prepare the seed PWB for scaling. Settings: settings.toml.
mode = "topology" also makes the seed match the GridView topology (this needs the EPCs).
"""
from pathlib import Path
from Scripts import pw, prep, settings, topo

cur_dir = Path(__file__).parent

# At import, not under __main__: multiprocessing workers re-import this module, and need the settings too.
run = settings.load()

# ------------------ Inputs ------------------
pw_dir = cur_dir / 'Seed'
gv_dir = cur_dir / 'HourEPCs'

# ------------------ Outputs ------------------
prepared_dir = cur_dir / 'Prepared'
fault_fp = prepared_dir / 'fault_duty.csv'
pvqv_fp = prepared_dir / 'pvqv.csv'
dummy_bus_fp = prepared_dir / 'DummyBus.aux'
log_fp = prepared_dir / 'Prepared_Log.xlsx'
created_elements_fp = prepared_dir / 'Prepared_CreatedElements.xlsx'

if(__name__=='__main__'):
    topology = run['mode'] == 'topology'
    print(f'Mode: {run["mode"]}')

    # First PowerWorld PWB. Topology mode: also the first GridView EPC (alphabetically), and all of them for XN loads.
    pw_fps = sorted(pw_dir.glob('*.pwb'))
    if len(pw_fps) == 0:
        raise SystemExit(f'No PowerWorld PWB found in {str(pw_dir)}')
    pw_fp = pw_fps[0]
    print(f'PW Input: {str(pw_fp)}')
    gv_fps = sorted(gv_dir.glob('*.epc'))
    if topology:
        if len(gv_fps) == 0:
            raise SystemExit(f'No GridView EPC found in {str(gv_dir)}')
        print(f'GV Input: {str(gv_fps[0])}')

    SimAuto = pw.connect()
    log = {}
    step = iter(range(1, 100))

    def save(step_name) -> Path:
        fp = prepared_dir / f'{next(step):02d}_{step_name}.pwb'
        pw.save_case(SimAuto, fp, run['case_format'])
        return fp

    def try_step(step_name, fn, required=False):
        """Runs fn; rolls it back if the case no longer solves. required: stop instead of going on without it."""
        with pw.trial(SimAuto) as t:
            result = fn()
        if not t.ok:
            print(f'WARNING: Did not solve after {step_name} !!! Rolled back.')
            if required:
                raise SystemExit(f'The case did not solve after {step_name}, so it was rolled back. Fix that, or set create_giant_swing = false.')
        return result

    if topology:
        print('create_dummy_bus_aux')
        pw.open_case(SimAuto, pw_fp)
        topo.create_dummy_bus_aux(SimAuto, dummy_bus_fp)

        print('create_missing_elements')
        topo.open_gv_case(SimAuto, gv_fps[0], dummy_bus_fp)
        gv_case_dict = topo.get_case_data(SimAuto)
        pw.open_case(SimAuto, pw_fp)
        pw_case_dict = topo.get_case_data(SimAuto)
        missing_dict = topo.create_missing_elements(SimAuto, gv_case_dict, pw_case_dict)
        pw.write_excel(created_elements_fp, missing_dict)
        save('create_missing_elements')

        print('fix_transformer_taps')
        log['bad_transformer_tap'] = topo.fix_transformer_taps(SimAuto)
        save('fix_transformer_taps')

        print('set_branch_statuses')
        pw_case_dict = topo.get_case_data(SimAuto)
        [log['branch_st_targets'], log['branch_st_change_failed']] = topo.set_branch_statuses(SimAuto, gv_case_dict, pw_case_dict)
        save('set_branch_statuses')
    else:
        pw.open_case(SimAuto, pw_fp)
        if not pw.solve(SimAuto):
            raise SystemExit(f'The seed case does not solve: {pw_fp}')

    if run['adjust_seed_shunts']:
        print('adjust_shunts')
        prep.adjust_shunts(SimAuto)
        save('adjust_shunts')

    if run['gen_terminal_voltage_control']:
        print('GenTerminalVoltageControl')
        try_step('GenTerminalVoltageControl.aux', lambda: pw.run(SimAuto, f'LoadAux("{cur_dir / "Scripts" / "GenTerminalVoltageControl.aux"}",YES);'))
        save('GenTerminalVoltageControl')

    prep.disable_area_agc(SimAuto)
    if not pw.solve(SimAuto):
        raise SystemExit('The case does not solve with Area AGC off. Fix that before step 02.')

    print('convert_loads_to_constant_power')
    log.update(prep.convert_loads_to_constant_power(SimAuto))
    save('convert_loads')

    if run['compute_pvqv']:
        print('get_pvqv')
        prep.get_pvqv(SimAuto).to_csv(pvqv_fp, index=False)

    if run['create_giant_swing']:
        swing_bus = run['swing_bus']
        if swing_bus == 'max_fault_duty':
            print('get_fault_duty')
            fault_df = prep.get_fault_duty(SimAuto)
            fault_df.to_csv(fault_fp, index=False)
            swing_bus = prep.max_fault_duty_bus(fault_df)
        elif swing_bus == 'slack':
            swing_bus = prep.slack_buses(SimAuto)[0]
        print(f'create_giant_swing on bus {swing_bus}')
        log['swing'] = try_step('create_giant_swing()', lambda: prep.create_giant_swing(SimAuto, swing_bus), required=True)
    fp = save('create_giant_swing' if run['create_giant_swing'] else 'disable_area_agc')

    if topology:
        print('create_distgen_XN_loads')
        gv_load_df = topo.get_gv_loads(SimAuto, gv_fps, dummy_bus_fp)
        pw.open_case(SimAuto, fp)
        log['distgen_loads'] = try_step('create_distgen_XN_loads()', lambda: topo.create_distgen_XN_loads(SimAuto, gv_load_df))

    pw.save_case(SimAuto, prepared_dir / 'Prepared.pwb', run['case_format'])

    print('Writing log.')
    if log:
        pw.write_excel(log_fp, log)

    SimAuto.CloseCase()
    SimAuto = None
    print('done')
