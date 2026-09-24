from pathlib import Path
import win32com.client
from Scripts import pw, topo

cur_dir = Path(__file__).parent

case_format = 'PWB23'

# ------------------ Inputs ------------------
# First GridView EPC, and first PowerWorld PWB.
gv_dir = cur_dir / 'HourEPCs'
gv_fps = sorted(gv_dir.glob("*.epc"))
if len(gv_fps) == 0:
    raise SystemExit(f'No GridView EPC found in {str(gv_dir)}')
gv_fp = gv_fps[0]
print(f'GV Input: {str(gv_fp)}')

pw_dir = cur_dir / 'Seed'
pw_fps = sorted(pw_dir.glob("*.pwb"))
if len(pw_fps) == 0:
    raise SystemExit(f'No PowerWorld PWB found in {str(pw_dir)}')
pw_fp = pw_fps[0]
print(f'PW Input: {str(pw_fp)}')

# ------------------ Outputs ------------------
toposeed_dir = cur_dir / 'TopoSeed'
fault_fp = toposeed_dir / 'fault_duty.csv'
pvqv_fp = toposeed_dir / 'pvqv.csv'
dummy_bus_fp = toposeed_dir / 'DummyBus.aux'
log_fp = toposeed_dir / 'TopoSeed_Log.xlsx'
created_elements_fp = toposeed_dir / 'TopoSeed_CreatedElements.xlsx'

if(__name__=='__main__'):
    SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")
    log = {}

    def save(step_name):
        pw.save_case(SimAuto, toposeed_dir / f'{step_name}.pwb', case_format)

    def try_step(step_name, fn):
        """Runs fn; rolls it back if the case no longer solves."""
        with pw.trial(SimAuto) as t:
            result = fn()
        if not t.ok:
            print(f'WARNING: Did not solve after {step_name} !!! Rolled back.')
        return result

    print('00_create_dummy_bus_aux')
    pw.open_case(SimAuto, pw_fp)
    topo.create_dummy_bus_aux(SimAuto, dummy_bus_fp)

    print('01_create_missing_elements')
    topo.open_gv_case(SimAuto, gv_fp, dummy_bus_fp)
    gv_case_dict = topo.get_case_data(SimAuto)
    pw.open_case(SimAuto, pw_fp)
    pw_case_dict = topo.get_case_data(SimAuto)
    missing_dict = topo.create_missing_elements(SimAuto, gv_case_dict, pw_case_dict)
    pw.write_excel(created_elements_fp, missing_dict)
    save('01_create_missing_elements')

    print('02_fix_transformer_taps')
    log['bad_transformer_tap'] = topo.fix_transformer_taps(SimAuto)
    save('02_fix_transformer_taps')

    print('03_set_branch_statuses')
    pw_case_dict = topo.get_case_data(SimAuto)
    [log['branch_st_targets'], log['branch_st_change_failed']] = topo.set_branch_statuses(SimAuto, gv_case_dict, pw_case_dict)
    save('03_set_branch_statuses')

    print('04_adjust_shunts')
    topo.adjust_shunts(SimAuto)
    save('04_adjust_shunts')

    print('05_GenTerminalVoltageControl')
    try_step('GenTerminalVoltageControl.aux', lambda: pw.run(SimAuto, f'LoadAux("{cur_dir / "Scripts" / "GenTerminalVoltageControl.aux"}",YES);'))
    save('05_GenTerminalVoltageControl')

    print('get_fault_duty')
    fault_df = topo.get_fault_duty(SimAuto)
    fault_df.to_csv(fault_fp, index=False)

    print('get_pvqv')
    topo.get_pvqv(SimAuto).to_csv(pvqv_fp, index=False)

    print('06_create_giant_swing')
    log['swing'] = try_step('create_giant_swing()', lambda: topo.create_giant_swing(SimAuto, fault_df))
    save('06_create_giant_swing')

    print('07_create_distgen_XN_loads')
    gv_load_df = topo.get_gv_loads(SimAuto, gv_fps, dummy_bus_fp)
    pw.open_case(SimAuto, toposeed_dir / '06_create_giant_swing.pwb')
    log['distgen_loads'] = try_step('create_distgen_XN_loads()', lambda: topo.create_distgen_XN_loads(SimAuto, gv_load_df))
    save('07_create_distgen_XN_loads')

    save('TopoSeed')

    print('Writing log.')
    pw.write_excel(log_fp, log)

    SimAuto.CloseCase()
    SimAuto = None
    print('done')
