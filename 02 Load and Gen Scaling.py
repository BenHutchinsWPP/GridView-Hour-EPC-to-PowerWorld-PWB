from pathlib import Path
import pandas as pd
import win32com.client
import Scripts.wpp_lib as wpp_lib
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

cur_dir = Path(__file__).parent

case_format = 'PWB23'

# ------------------ Inputs ------------------
gv_dir = cur_dir / 'HourEPCs'
pw_fp = cur_dir / 'TopoSeed' / 'TopoSeed.pwb'

# ------------------ Outputs ------------------
fault_fp = cur_dir / 'TopoSeed' / 'fault_duty.csv'
pvqv_fp = cur_dir / 'TopoSeed' / 'pvqv.csv'
toposeed_log_fp = cur_dir / 'TopoSeed' / 'TopoSeed_Log.xlsx'

def create_case(SimAuto, gv_fp, pw_fp):
    target_fp = cur_dir / 'Output' / (gv_fp.stem + '_01_Target.xlsx')
    target_test_fp = cur_dir / 'Output' / (gv_fp.stem + '_02_TargetTest.xlsx')
    scale_log_fp = cur_dir / 'Output' / (gv_fp.stem + '_03_ScaleLog.xlsx')

    print('compute_pw_targets')
    [gen_target_df, load_target_df] = wpp_lib.compute_pw_targets(SimAuto, gv_fp, pw_fp)
    wpp_lib.df_dict_to_excel_workbook(target_fp, {
        'gen':gen_target_df
        ,'load':load_target_df
    })

    print('test_gen_targets_parallel')
    gen_target_df = wpp_lib.test_gen_targets_parallel(pw_fp, gen_target_df)
    wpp_lib.df_dict_to_excel_workbook(target_test_fp, {
        'gen':gen_target_df
    })
    
    # Exclude generation changes which do not solve successfully on their own. 
    gen_target_df.loc[gen_target_df['Success'] == False, ['Include', 'ExclusionReason']] = [False, 'Individual Gen Test Diverged']

    # Don't adjust the swing unit. 
    print('get_swing')
    swing_df = pd.read_excel(toposeed_log_fp, sheet_name='swing')
    swing_bus = swing_df.loc[swing_df.index[0], 'BusNum']
    swing_id = swing_df.loc[swing_df.index[0], 'ID']
    gen_target_df.loc[
        (gen_target_df['BusNum'] == swing_bus) & (gen_target_df['ID'] == swing_id), 
        ['Status_Target', 'Include', 'ExclusionReason']
    ] = ['Closed', False, 'Swing Unit']
    
    print('get_pvqv_csv')
    pvqv_df = pd.read_csv(pvqv_fp)

    print('iterate_to_gen_load_targets')
    wpp_lib.report_gen_load_balance(gen_target_df, load_target_df)
    if not wpp_lib.open_case(SimAuto, pw_fp):
        raise
    scalelog_dict = wpp_lib.iterate_to_gen_load_targets(SimAuto, gen_target_df, load_target_df, pvqv_df)
    wpp_lib.df_dict_to_excel_workbook(scale_log_fp, scalelog_dict)
    wpp_lib.save_case(SimAuto, cur_dir / 'Output' / (gv_fp.stem + '.pwb'),case_format)

    # Exit. 
    SimAuto.CloseCase()
    return

if(__name__=='__main__'):
    for gv_fp in gv_dir.glob('*.epc'):
        gv_fp = Path(gv_fp)
        create_case(SimAuto, gv_fp, pw_fp)

    SimAuto = None
    print('done')


