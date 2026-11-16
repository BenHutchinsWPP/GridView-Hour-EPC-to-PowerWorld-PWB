"""
Step 02: scale the prepared seed to each GridView hour EPC in ./HourEPCs/. Settings: settings.toml.
mode = "area": per-area load and per-area & GenType generation (Scripts/area.py).
mode = "topology": every gen & load to its own EPC value (topo.build_targets()).
"""
from pathlib import Path
import traceback
import pandas as pd
from Scripts import area, prep, pw, scale, settings, topo

cur_dir = Path(__file__).parent

# At import, not under __main__: multiprocessing workers re-import this module, and need the settings too.
run = settings.load()

# ------------------ Inputs ------------------
gv_dir = cur_dir / 'HourEPCs'
pw_fp = cur_dir / 'Prepared' / 'Prepared.pwb'
pvqv_fp = cur_dir / 'Prepared' / 'pvqv.csv'
dummy_bus_fp = cur_dir / 'Prepared' / 'DummyBus.aux'

# ------------------ Outputs ------------------
out_dir = cur_dir / 'Output'

def create_case(SimAuto, gv_fp, out_fp):
    print('targets')
    if run['mode'] == 'topology':
        targets, reports = topo.compute_pw_targets(SimAuto, gv_fp, pw_fp, dummy_bus_fp), {}
    else:
        targets, reports = area.compute_targets(SimAuto, gv_fp, pw_fp)
    pw.write_excel(out_dir / (gv_fp.stem + '_01_Target.xlsx'), {**reports, **{table.lower(): df for table, df in targets.items()}})

    if run['pretest_gens']:
        print('test_gen_targets_parallel')
        gen_df = targets['Gen'] = scale.test_gen_targets_parallel(pw_fp, targets['Gen'])
        pw.write_excel(out_dir / (gv_fp.stem + '_02_TargetTest.xlsx'), {'gen': gen_df})
        # Exclude generation changes which do not solve successfully on their own.
        scale.exclude(gen_df, ~gen_df['Success'], 'Individual Gen Test Diverged')

    print('iterate_to_gen_load_targets')
    scale.report_gen_load_balance(targets)
    pw.open_case(SimAuto, pw_fp)
    if run['mode'] == 'area':
        # Open loads aren't scaled, but their distributed generation is set Open & 0 too.
        case_load = pw.get_df(SimAuto, 'Load', area.LOAD_FIELDS)
        pw.set_df(SimAuto, 'Load', area.open_load_cleanup(case_load), ['BusNum', 'ID', 'DistStatus', 'DistMWInput', 'DistMvarInput'])
    pvqv = pd.read_csv(pvqv_fp) if run['use_pvqv'] else None
    if run['balance_swing']:
        # Balance the swing whenever it drifts far during scaling, and again at the end.
        after_step, midrun_log, midrun_sheets = area.midrun_balancer(SimAuto)
        log = scale.iterate_to_gen_load_targets(SimAuto, targets, pvqv, after_step)
        log['midrun_balance_df'] = pd.DataFrame(midrun_log)
        log.update(midrun_sheets)
        log.update(area.balance_swing(SimAuto, targets))
    else:
        log = scale.iterate_to_gen_load_targets(SimAuto, targets, pvqv)
    if reports:
        # Read back from PowerWorld, so the report shows what the case holds rather than what the scaling meant to do.
        log = {**area.achieved(area.read_case(SimAuto), reports, prep.slack_buses(SimAuto)), **log}
    if not pw.solve(SimAuto):
        print('WARNING: the case does not solve at the end of scaling. Saved as it is.')
    # Save the case first: a log which can't be written (e.g. open in Excel) mustn't cost the solved hour.
    pw.save_case(SimAuto, out_fp, run['case_format'])
    log_fp = out_dir / (gv_fp.stem + '_03_ScaleLog.xlsx')
    try:
        pw.write_excel(log_fp, scale.merge_network_changes(log))
    except OSError as e:
        print(f'WARNING: could not write {log_fp.name} (open in Excel?): {e!r}. The case was saved; '
              f'delete {out_fp.name} to rerun this hour with its log.')

    SimAuto.CloseCase()

if(__name__=='__main__'):
    print(f'Mode: {run["mode"]}')
    if run['use_pvqv'] and not pvqv_fp.exists():
        raise SystemExit(f'use_pvqv needs {pvqv_fp}: run step 01 with compute_pvqv = true.')
    # Created here, not at import: multiprocessing workers re-import this module.
    SimAuto = pw.connect()
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
