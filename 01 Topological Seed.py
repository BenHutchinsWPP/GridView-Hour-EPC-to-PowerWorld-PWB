from pathlib import Path
import pandas as pd
import win32com.client
import Scripts.wpp_lib as wpp_lib
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

cur_dir = Path(__file__).parent

case_format = 'PWB23'

# ------------------ Inputs ------------------
# Get first Gridview EPC
gv_dir = cur_dir / 'HourEPCs'
gv_fps = list(gv_dir.glob("*.epc"))
if len(gv_fps) == 0:
    print(f'No GridView EPC found in {str(gv_dir)}')
    exit()
gv_fp = gv_fps[0]
print(f'GV Input: {str(gv_fp)}')

# Get first PowerWorld PWB
pw_dir = cur_dir / 'Seed'
pw_fps = list(pw_dir.glob("*.pwb"))
if len(pw_fps) == 0:
    print(f'No PowerWorld PWB found in {str(pw_dir)}')
    exit()
pw_fp = pw_fps[0]
print(f'PW Input: {str(pw_fp)}')

# ------------------ Outputs ------------------
fault_fp = cur_dir / 'TopoSeed' / 'fault_duty.csv'
pvqv_fp = cur_dir / 'TopoSeed' / 'pvqv.csv'
dummy_bus_fp = cur_dir / 'TopoSeed' / 'DummyBus.aux'
errors_fp = cur_dir / 'TopoSeed' / 'TopoSeed_Log.xlsx'
created_elements_fp = cur_dir / 'TopoSeed' / 'TopoSeed_CreatedElements.xlsx'

if(__name__=='__main__'):
    print('Initializing log.')
    writer = pd.ExcelWriter(errors_fp, engine='openpyxl')

    print('00_create_dummy_bus_aux')
    if not wpp_lib.open_case(SimAuto, pw_fp):
        quit()
    wpp_lib.create_dummy_bus_aux(SimAuto, dummy_bus_fp)

    print('01_create_missing_elements')
    if not wpp_lib.open_case(SimAuto, gv_fp):
        quit()
    # Renumber dummy buses before getting case data
    retVal = SimAuto.RunScriptCommand('EnterMode(EDIT);')
    retVal = SimAuto.RunScriptCommand('LoadAux("'+str(dummy_bus_fp)+'",YES);')
    print(retVal)
    gv_case_dict = wpp_lib.get_case_data(SimAuto)

    if not wpp_lib.open_case(SimAuto, pw_fp):
        quit()
    pw_case_dict = wpp_lib.get_case_data(SimAuto)
    missing_dict = wpp_lib.create_missing_elements(SimAuto, gv_case_dict, pw_case_dict)
    wpp_lib.df_dict_to_excel_workbook(created_elements_fp, missing_dict)
    wpp_lib.save_case(SimAuto, cur_dir / 'TopoSeed' / '01_create_missing_elements.pwb', case_format)

    print('02_fix_transformer_taps')
    pw_case_dict = wpp_lib.get_case_data(SimAuto)
    bad_transformer_df = wpp_lib.fix_transformer_taps(SimAuto)
    bad_transformer_df.to_excel(writer, sheet_name='bad_transformer_tap', index=False)
    wpp_lib.save_case(SimAuto, cur_dir / 'TopoSeed' / '02_fix_transformer_taps.pwb', case_format)

    print('03_set_branch_statuses')
    pw_case_dict = wpp_lib.get_case_data(SimAuto)
    [status_targets_df, fail_df] = wpp_lib.set_branch_statuses(SimAuto, gv_case_dict, pw_case_dict)
    fail_df.to_excel(writer, sheet_name='branch_st_change_failed', index=False)
    status_targets_df.to_excel(writer, sheet_name='branch_st_targets', index=False)
    wpp_lib.save_case(SimAuto, cur_dir / 'TopoSeed' / '03_set_branch_statuses.pwb', case_format)

    print('04_adjust_shunts')
    wpp_lib.adjust_shunts(SimAuto)
    wpp_lib.save_case(SimAuto, cur_dir / 'TopoSeed' / '04_adjust_shunts.pwb', case_format)

    print('05_GenTerminalVoltageControl')
    SimAuto.SaveState()
    retVal = SimAuto.RunScriptCommand('SetCurrentDirectory("'+str(cur_dir)+'");')
    retVal = SimAuto.RunScriptCommand('LoadAux("Scripts/GenTerminalVoltageControl.aux",YES);')
    print(retVal)
    if not wpp_lib.solve(SimAuto, wpp_lib.mva_mismatch_threshold):
        print('WARNING: Did not solve after running GenTerminalVoltageControl.aux !!!')
        SimAuto.LoadState()
    wpp_lib.save_case(SimAuto, cur_dir / 'TopoSeed' / '05_GenTerminalVoltageControl.pwb', case_format)

    print('get_fault_duty')
    fault_df = wpp_lib.get_fault_duty(SimAuto)
    fault_df.to_csv(fault_fp, index=False)

    print('get_pvqv')
    pvqv_df = wpp_lib.get_pvqv(SimAuto)
    pvqv_df.to_csv(pvqv_fp, index=False)
    
    print('06_create_giant_swing')
    SimAuto.SaveState()
    swing_df = wpp_lib.create_giant_swing(SimAuto, fault_df)
    swing_df.to_excel(writer, sheet_name='swing', index=False)
    if not wpp_lib.solve(SimAuto, wpp_lib.mva_mismatch_threshold):
        print('WARNING: Did not solve after running create_giant_swing() !!!')
        SimAuto.LoadState()
    case_fp = cur_dir / 'TopoSeed' / '06_create_giant_swing.pwb'
    wpp_lib.save_case(SimAuto, case_fp, case_format)

    print('07_create_distgen_XN_loads')
    SimAuto.SaveState()
    distgen_loads_df = wpp_lib.create_distgen_XN_loads(SimAuto, gv_fps, case_fp)
    distgen_loads_df.to_excel(writer, sheet_name='distgen_loads', index=False)
    if not wpp_lib.solve(SimAuto, wpp_lib.mva_mismatch_threshold):
        print('WARNING: Did not solve after running create_distgen_XN_loads() !!!')
        SimAuto.LoadState()
    wpp_lib.save_case(SimAuto, cur_dir / 'TopoSeed' / '07_create_distgen_XN_loads.pwb', case_format)

    wpp_lib.save_case(SimAuto, cur_dir / 'TopoSeed' / 'TopoSeed.pwb', case_format)

    print('Cleaning up before exit.')
    try:
        writer.close()
    except:
        pass

    SimAuto.CloseCase()
    SimAuto = None
    print('done')


