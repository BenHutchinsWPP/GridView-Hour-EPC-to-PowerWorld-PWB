from pathlib import Path
import traceback
import pandas as pd
import win32com.client
from Scripts import pw, topo, scale

cur_dir = Path(__file__).parent

case_format = 'PWB23'

# ------------------ Options ------------------
# How to handle gens/loads which make the case diverge:
#   'pretest': Test each gen target on its own first (in parallel), and exclude those which diverge.
#              Scaling stops at the first step which diverges.
#   'bisect':  No pretest. When a scaling step diverges, bisect it to find the gens/loads which diverge,
#              exclude them at their last good step, and continue scaling the rest.
divergence_mode = 'pretest'

# ------------------ Inputs ------------------
gv_dir = cur_dir / 'HourEPCs'
pw_fp = cur_dir / 'TopoSeed' / 'TopoSeed.pwb'
pvqv_fp = cur_dir / 'TopoSeed' / 'pvqv.csv'
dummy_bus_fp = cur_dir / 'TopoSeed' / 'DummyBus.aux'

# ------------------ Outputs ------------------
out_dir = cur_dir / 'Output'

def create_case(SimAuto, gv_fp, out_fp):
    print('compute_pw_targets')
    targets = scale.compute_pw_targets(SimAuto, gv_fp, pw_fp, dummy_bus_fp)
    pw.write_excel(out_dir / (gv_fp.stem + '_01_Target.xlsx'), {table.lower(): df for table, df in targets.items()})

    # Don't adjust the swing unit.
    gen_df = targets['Gen']
    swing = gen_df['ID'] == topo.SWING_ID
    gen_df.loc[swing, 'Status_Target'] = 'Closed'
    scale.exclude(gen_df, swing, 'Swing Unit')

    if divergence_mode == 'pretest':
        print('test_gen_targets_parallel')
        gen_df = targets['Gen'] = scale.test_gen_targets_parallel(pw_fp, gen_df)
        pw.write_excel(out_dir / (gv_fp.stem + '_02_TargetTest.xlsx'), {'gen': gen_df})
        # Exclude generation changes which do not solve successfully on their own.
        scale.exclude(gen_df, ~gen_df['Success'], 'Individual Gen Test Diverged')

    print('iterate_to_gen_load_targets')
    scale.report_gen_load_balance(targets)
    pw.open_case(SimAuto, pw_fp)
    scalelog_dict = scale.iterate_to_gen_load_targets(SimAuto, targets, pd.read_csv(pvqv_fp), bisect=divergence_mode == 'bisect')
    pw.write_excel(out_dir / (gv_fp.stem + '_03_ScaleLog.xlsx'), scalelog_dict)
    pw.save_case(SimAuto, out_fp, case_format)

    SimAuto.CloseCase()

if(__name__=='__main__'):
    if divergence_mode not in ('pretest', 'bisect'):
        raise SystemExit(f"divergence_mode must be 'pretest' or 'bisect', not {divergence_mode!r}")

    # Created here, not at import: multiprocessing workers re-import this module.
    SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")
    for gv_fp in sorted(gv_dir.glob('*.epc')):
        out_fp = out_dir / (gv_fp.stem + '.pwb')
        if out_fp.exists():
            print(f'Skipping {gv_fp.name}: {out_fp.name} already exists. Delete it to re-run.')
            continue
        # Keep going with the remaining hours if one fails.
        try:
            create_case(SimAuto, gv_fp, out_fp)
        except Exception:
            print(f'FAILED: {gv_fp.name}')
            traceback.print_exc()

    SimAuto = None
    print('done')
