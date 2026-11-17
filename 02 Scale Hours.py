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
# Quiet the three opens an hour (step 01 keeps the line).
pw.print_opened = pw.print_solve_errors = False

# ------------------ Inputs ------------------
gv_dir = cur_dir / 'HourEPCs'
pw_fp = cur_dir / 'Prepared' / 'Prepared.pwb'
pvqv_fp = cur_dir / 'Prepared' / 'pvqv.csv'
dummy_bus_fp = cur_dir / 'Prepared' / 'DummyBus.aux'

# ------------------ Outputs ------------------
out_dir = cur_dir / 'Output'

def summary_line(hour, log, area_mode):
    """The hour's one-line result, from the sheets of its log."""
    runs, exc = log['runs'], log['exceptions']
    swing = runs['Swing MW After'].iloc[-1] if len(runs) else float('nan')
    n_statcom, n_branch = int((exc['Table'] == 'STATCOM').sum()), int((exc['Table'] == 'Branch').sum())
    n_exc = len(exc) - n_statcom - n_branch
    line = f'{hour}: swing {swing:.1f} MW; {n_exc} exceptions; {n_statcom} STATCOMs, {n_branch} branches opened'
    if area_mode:
        err = log['areas']['Error MW'].abs()
        if err.notna().any():
            worst = log['areas'].loc[err.idxmax()]
            line += f'; worst error area {worst["AreaNumber"]} {worst["Group"]} {worst["Error MW"]:.0f} MW'
    return line

def clear_dist_gen(SimAuto, swing) -> list[scale.Run]:
    """
    GridView ignores distributed generation: sets it Open & 0 on every load still carrying it (loads the scaling never
    touched or didn't finish), one solve-checked edit, bisected if it diverges. The final balance (if balance_swing)
    takes up the MW. Returns a `distgen` run for the runs sheet (none if nothing was left), whose Note names the loads
    which kept theirs.
    """
    left = area.dist_gen_cleanup(pw.get_df(SimAuto, 'Load', area.LOAD_FIELDS))
    if not len(left):
        return []
    before = swing()
    failed = pw.set_df_bisect(SimAuto, 'Load', left, ['BusNum', 'ID', 'DistStatus', 'DistMWInput', 'DistMvarInput'])
    done = left.drop(index=failed)
    kept = (f'{len(failed)} kept theirs (the case diverged without it): ' + ', '.join(left.loc[failed, 'ObjectID'].head(5))
            + (' ...' if len(failed) > 5 else '')) if failed else ''
    note = '; '.join(filter(None, [f'dist gen cleared on {len(done)} loads ({area.dist_gen_removed(done):,.0f} MW)', kept]))
    print(note[0].upper() + note[1:] + '.')
    return [scale.Run('distgen', before=before, after=swing(), note=note)]

def create_case(SimAuto, gv_fp, out_fp):
    area_mode = run['mode'] != 'topology'
    if area_mode:
        targets, reports = area.compute_targets(SimAuto, gv_fp, pw_fp)
    else:
        targets, reports = topo.compute_pw_targets(SimAuto, gv_fp, pw_fp, dummy_bus_fp), {}
    pretest = []
    if run['pretest_gens']:
        gen_df = targets['Gen'] = scale.test_gen_targets_parallel(pw_fp, targets['Gen'])
        # Exclude generation changes which do not solve successfully on their own.
        scale.exclude(gen_df, ~gen_df['Success'], 'Individual Gen Test Diverged')
        # Only for the plan written below: the main run's log lists these rows again, once scaling has run.
        pretest = [scale.Run('pretest', {'gen': gen_df[~gen_df['Success']]})]
    scale.report_gen_load_balance(targets)
    pre_runs = []
    if run['balance_swing']:
        # Fold the targets' own imbalance in now, so the swing doesn't drift (and set off balances) as they scale.
        losses = reports.get('area_loss')       # Area mode only.
        loss_change = float((losses['Est. Loss MW'] - losses['Seed Loss MW']).sum()) if losses is not None else 0.0
        if losses is not None and len(losses):
            print(f'Losses: seed {losses["Seed Loss MW"].sum():,.0f} MW, estimated {losses["Est. Loss MW"].sum():,.0f} MW '
                  f'at the target load (loss_exponent {area.LOSS_EXPONENT:g}).')
        moved = area.prebalance_targets(targets, loss_change)
        if moved is not None:
            pre_runs = [scale.Run('prebalance', moved=moved, note='targets moved before scaling')]
    log_fp = out_dir / (gv_fp.stem + '_Log.xlsx')

    def build_log(runs, final, balance_tables=(), swing_buses=(), exception_runs=(), final_loss=None):
        log = {'exceptions': scale.exceptions([*exception_runs, *runs], targets, final), 'runs': scale.runs_sheet(runs)}
        if area_mode:
            # Before scaling (final None) nothing has been reached or balanced yet: the plan alone.
            scaled = targets if final is not None else None
            log = {'areas': area.areas(final, reports, scaled, balance_tables, swing_buses, final_loss), **log}
        return log

    # The plan first, so that a failure during the hour's long scaling still leaves it.
    try:
        # Units left out as outside the main island, listed as the main run's log will list them.
        islanded = scale.Run('main', {t.lower(): df[scale.is_island_reason(df['ExclusionReason'])] for t, df in targets.items()})
        pw.write_excel(log_fp, build_log([], None, exception_runs=[*pretest, islanded]))
    except Exception as e:     # Never lose the hour over its log.
        print(f'WARNING: could not write {log_fp.name} (open in Excel?): {e!r}.')

    pw.open_case(SimAuto, pw_fp)
    scale.turn_off_islanded(SimAuto, targets)     # Everything outside the main island goes Open, before scaling.
    pvqv = pd.read_csv(pvqv_fp) if run['use_pvqv'] else None
    balance_tables = []
    slack = prep.slack_buses(SimAuto)
    swing = lambda: area.swing_mw(pw.get_df(SimAuto, 'Gen', area.GEN_FIELDS), slack)[0]
    before = swing()
    midrun_runs = []
    if run['balance_swing']:
        # Balance the swing whenever it drifts far during scaling, and again at the end.
        after_step, midrun_runs = area.midrun_balancer(SimAuto)
        log = scale.iterate_to_gen_load_targets(SimAuto, targets, pvqv, after_step)
    else:
        log = scale.iterate_to_gen_load_targets(SimAuto, targets, pvqv)
    # Main last of the runs during scaling (its midrun balances happened inside it), so its After is where scaling left
    # the swing; the final balance rounds follow. The last row's After is where the swing ended.
    runs = [*pre_runs, *midrun_runs, scale.Run('main', log, before=before, after=swing())]
    runs += clear_dist_gen(SimAuto, swing)
    if run['balance_swing']:
        runs += area.balance_swing(SimAuto, targets, tables=balance_tables)
    # Read back from PowerWorld, so the log shows what the case holds rather than what the scaling meant to do.
    final = area.read_case(SimAuto) if area_mode else {t: pw.get_df(SimAuto, t, f) for t, f in topo.TARGET_FIELDS.items()}
    final_loss = area.area_losses(SimAuto) if area_mode else None      # The areas sheet's Final Loss MW.
    swing_buses = prep.slack_buses(SimAuto)
    if not pw.solve(SimAuto):
        print('WARNING: the case does not solve at the end of scaling. Saved as it is.')
    # Save the case first: a log which can't be built or written (e.g. open in Excel) mustn't cost the solved hour.
    pw.save_case(SimAuto, out_fp, run['case_format'])
    try:
        log = build_log(runs, final, balance_tables, swing_buses, final_loss=final_loss)
        print(summary_line(gv_fp.stem, log, area_mode))
        pw.write_excel(log_fp, log)
    except Exception as e:     # Never lose the hour over its log.
        print(f'WARNING: could not write {log_fp.name} (open in Excel?): {e!r}. The case was saved; '
              f'delete {out_fp.name} to rerun this hour with its log.')

    SimAuto.CloseCase()

if(__name__=='__main__'):
    print(f'Mode: {run["mode"]}')
    if run['use_pvqv'] and not pvqv_fp.exists():
        raise SystemExit(f'use_pvqv needs {pvqv_fp}: run step 01 with compute_pvqv = true.')
    # Created here, not at import: multiprocessing workers re-import this module.
    SimAuto = pw.connect()
    gv_fps = sorted(gv_dir.glob('*.epc'))
    for n, gv_fp in enumerate(gv_fps, 1):
        out_fp = out_dir / (gv_fp.stem + '.pwb')
        if out_fp.exists():
            print(f'Skipping {gv_fp.name}: {out_fp.name} already exists. Delete it to re-run.')
            continue
        print(f'=== {gv_fp.name} ({n} of {len(gv_fps)}) ===')
        # Keep going with the remaining hours if one fails.
        try:
            create_case(SimAuto, gv_fp, out_fp)
        except Exception:
            print(f'FAILED: {gv_fp.name}')
            traceback.print_exc()

    SimAuto = None
    print('done')
