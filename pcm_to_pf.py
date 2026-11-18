"""
The whole run, settings in settings.toml:
  step 01, prepare: the seed PWB (./Seed/) -> ./Prepared/Prepared.pwb. Skipped while Prepared.pwb was made from the
           same settings and seed (Prepared_Settings.json): change the mode, say, and it is redone.
  step 02, scale: Prepared.pwb to each GridView hour EPC in ./HourEPCs/ -> ./Output/<hour>.pwb and <hour>_Log.xlsx.
           [run] workers = N scales N hours at once, each in its own PowerWorld instance (console: Output/<hour>_Console.txt).
  step 03, merge: every Output/<hour>_Log.xlsx -> '03 Merge Reports.xlsx'.
With no arguments it runs the [run] steps (all three by default). `python pcm_to_pf.py scale merge` runs those steps
only. Parallel scaling is set in settings.toml only: [run] workers.
mode = "area": per-area load and per-area & GenType generation (Scripts/area.py).
mode = "topology": step 01 also makes the seed match the GridView topology (needs the EPCs); step 02 sets every gen &
load to its own EPC value (topo.build_targets()).
"""
import argparse
import contextlib
import json
import multiprocessing as mp
import multiprocessing.connection
import os
import re
from pathlib import Path
import traceback
import numpy as np
import pandas as pd
from Scripts import area, prep, pw, scale, settings, topo

cur_dir = Path(__file__).parent

# At import, not under __main__: the worker processes (pretest_gens, workers > 1) re-import this module, and need the
# settings too.
run = settings.load()

# ------------------ Folders and files ------------------
seed_dir = cur_dir / 'Seed'                 # Step 01 input: the first PWB in it.
gv_dir = cur_dir / 'HourEPCs'               # The hour EPCs.
prepared_dir = cur_dir / 'Prepared'         # Step 01 output, step 02 input.
pw_fp = prepared_dir / 'Prepared.pwb'
pvqv_fp = prepared_dir / 'pvqv.csv'
dummy_bus_fp = prepared_dir / 'DummyBus.aux'
fault_fp = prepared_dir / 'fault_duty.csv'
prep_log_fp = prepared_dir / 'Prepared_Log.xlsx'
created_elements_fp = prepared_dir / 'Prepared_CreatedElements.xlsx'
prep_stamp_fp = prepared_dir / 'Prepared_Settings.json'   # What Prepared.pwb was made from (prep_stamp()).
out_dir = cur_dir / 'Output'                # Step 02 output, step 03 input.

STEPS = settings.STEPS      # ('prepare', 'scale', 'merge')

# ================== Step 01: prepare the seed ==================
# The [run] settings step 01 reads. With [prep], [topology] and [pw], the seed file and (topology mode) the EPCs, they
# are what Prepared.pwb depends on.
PREP_RUN_KEYS = ('mode', 'case_format', 'create_giant_swing', 'swing_bus', 'adjust_seed_shunts',
                 'gen_terminal_voltage_control', 'compute_pvqv')

def prep_stamp() -> dict:
    """What step 01 would make Prepared.pwb from now, as plain JSON values."""
    stamp = {k: run[k] for k in PREP_RUN_KEYS}
    for section in ('prep', 'topology', 'pw'):
        module = settings.MODULES[section]
        stamp[section] = {k: getattr(module, k.upper()) for k in sorted(settings.SETTINGS[section])}
    seed = sorted(seed_dir.glob('*.pwb'))
    stamp['seed'] = [seed[0].name, seed[0].stat().st_size, seed[0].stat().st_mtime_ns] if seed else None
    if run['mode'] == 'topology':       # The topology and the XN loads come from the EPCs.
        stamp['epcs'] = [fp.name for fp in sorted(gv_dir.glob('*.epc'))]
    return json.loads(json.dumps(stamp, default=sorted))       # Sets as sorted lists, tuples as lists.

def prepared_changes() -> list[str]:
    """Why Prepared.pwb doesn't match settings.toml and its inputs now; [] if it does, or if there is none."""
    if not pw_fp.exists():
        return []
    try:
        old = json.loads(prep_stamp_fp.read_text(encoding='utf-8'))
    except FileNotFoundError:
        return [f'no {prep_stamp_fp.name} records what it was made from']
    except ValueError:
        return [f'{prep_stamp_fp.name} is unreadable']
    new = prep_stamp()
    return [f'{k}: {old.get(k)!r} -> {new.get(k)!r}' for k in sorted(set(old) | set(new)) if old.get(k) != new.get(k)]

def prepare_if_needed():
    """Step 01, unless Prepared.pwb was made from the same settings and inputs."""
    changes = prepared_changes()
    if pw_fp.exists() and not changes:
        print(f'Skipping step 01: {pw_fp.name} matches settings.toml and the seed. Delete it to re-prepare anyway.')
        return
    if changes:
        print(f'Redoing step 01: {pw_fp.name} was made from other settings or inputs: ' + '; '.join(changes))
        done = sorted(out_dir.glob('*.pwb'))
        if done:
            print(f'WARNING: {len(done)} hours in {out_dir} were scaled from the old {pw_fp.name}, and step 02 skips '
                  'them: delete them to rescale.')
    # Nothing stale is left behind if the preparation fails.
    pw_fp.unlink(missing_ok=True)
    prep_stamp_fp.unlink(missing_ok=True)
    prepare()

def prepare():
    stamp = prep_stamp()        # Before it runs: a seed changed meanwhile shows up as stale next time.
    topology = run['mode'] == 'topology'
    # First PowerWorld PWB. Topology mode: also the first GridView EPC (alphabetically), and all of them for XN loads.
    seed_fps = sorted(seed_dir.glob('*.pwb'))
    if len(seed_fps) == 0:
        raise SystemExit(f'No PowerWorld PWB found in {str(seed_dir)}')
    seed_fp = seed_fps[0]
    print(f'PW Input: {str(seed_fp)}')
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
        pw.open_case(SimAuto, seed_fp)
        topo.create_dummy_bus_aux(SimAuto, dummy_bus_fp)

        print('create_missing_elements')
        topo.open_gv_case(SimAuto, gv_fps[0], dummy_bus_fp)
        gv_case_dict = topo.get_case_data(SimAuto)
        pw.open_case(SimAuto, seed_fp)
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
        pw.open_case(SimAuto, seed_fp)
        if not pw.solve(SimAuto):
            raise SystemExit(f'The seed case does not solve: {seed_fp}')

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

    pw.save_case(SimAuto, pw_fp, run['case_format'])
    prep_stamp_fp.write_text(json.dumps(stamp, indent=2), encoding='utf-8')

    print('Writing log.')
    if log:
        pw.write_excel(prep_log_fp, log)

    SimAuto.CloseCase()

# ================== Step 02: scale to each hour ==================
def summary_line(hour, log) -> str:
    """The hour's one-line result, from the same `hours` row step 03 merges (_hour_row())."""
    row = _hour_row(hour, log)
    line = (f'{hour}: swing {row["Final Swing MW"]:.1f} MW; {row["Exceptions Gen/Load"]} exceptions; '
            f'{row["Exceptions STATCOM"]} STATCOMs, {row["Exceptions Branch"]} branches opened')
    if pd.notna(row['Worst Error MW']):        # Area mode: the areas sheet's worst |Error MW|.
        line += f'; worst error area {row["Worst AreaNumber"]} {row["Worst Group"]} {row["Worst Error MW"]:.0f} MW'
    return line

def clear_dist_gen(SimAuto, swing) -> list[scale.Run]:
    """
    Safety net: GridView ignores distributed generation, so sets it Open & 0 on every load still carrying it. On a seed
    converted by step 01 that finds only loads whose conversion (or dist gen clearing) diverged there. One solve-checked
    edit, bisected if it diverges. The final balance (if balance_swing) takes up the MW. Returns a `distgen` run for the
    runs sheet (none if nothing was left), whose Note names the loads which kept theirs.
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

def create_case(SimAuto, gv_fp, out_fp) -> str | None:
    """Scales one hour. Returns its summary line (None if its log couldn't be built)."""
    pw.print_opened = pw.print_solve_errors = False     # Quiet the three opens an hour (step 01 keeps the line).
    area_mode = run['mode'] != 'topology'
    if area_mode:
        targets, reports = area.compute_targets(SimAuto, gv_fp, pw_fp)
    else:
        targets, reports = topo.compute_pw_targets(SimAuto, gv_fp, pw_fp, dummy_bus_fp)
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
        # The case's own slack gens whose key is in the EPC end at their EPC MW: the other gens make up their move.
        swing_change = float((reports['swing']['Seed MW'] - reports['swing']['EPC MW']).sum())
        moved = area.prebalance_targets(targets, loss_change, swing_change=swing_change)
        if moved is not None:
            pre_runs = [scale.Run('prebalance', moved=moved, note='targets moved before scaling')]
    log_fp = out_dir / (gv_fp.stem + '_Log.xlsx')

    def build_log(runs, final, swing_buses=(), exception_runs=(), final_loss=None):
        log = {'exceptions': scale.exceptions([*exception_runs, *runs], targets, final), 'runs': scale.runs_sheet(runs)}
        if area_mode:
            # Before scaling (final None) nothing has been reached or balanced yet: the plan alone.
            scaled = targets if final is not None else None
            log = {'areas': area.areas(final, reports, scaled, runs, swing_buses, final_loss), **log, 'shortfall': reports['shortfall']}
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
    try:        # An old Prepared.pwb: rerun step 01. Only a warning, so it never costs the hour.
        area.warn_unconverted_loads(pw.get_df(SimAuto, 'Load', area.LOAD_FIELDS))
    except Exception as e:
        print(f'WARNING: could not check the loads for current/impedance/dist gen: {e!r}.')
    pvqv = pd.read_csv(pvqv_fp) if run['use_pvqv'] else None
    slack = prep.slack_buses(SimAuto)
    swing = lambda: area.swing_mw(pw.get_df(SimAuto, 'Gen', area.GEN_FIELDS), slack)[0]
    before = swing()
    # Balance the swing whenever it drifts far during scaling, and again at the end.
    after_step, midrun_runs = area.midrun_balancer(SimAuto, reports['swing']) if run['balance_swing'] else (None, [])
    log = scale.iterate_to_gen_load_targets(SimAuto, targets, pvqv, after_step)
    # Main last of the runs during scaling (its midrun balances happened inside it), so its After is where scaling left
    # the swing; the final balance rounds follow. The last row's After is where the swing ended.
    runs = [*pre_runs, *midrun_runs, scale.Run('main', log, before=before, after=swing())]
    try:
        runs += clear_dist_gen(SimAuto, swing)
    except Exception as e:     # Never lose the hour over its dist gen: it stays on the loads which kept it.
        print(f'WARNING: clearing dist gen failed: {e!r}. Carrying on without it.')
        runs.append(scale.Run('distgen', note=f'failed: {e!r}'))
    if area_mode and area.MAKEUP:
        # What excluded units didn't reach goes onto their group's other units, before the final balance.
        try:
            runs += area.makeup(SimAuto, targets, nonconforming=reports['nonconforming'])
        except Exception as e:     # Never lose the hour over its make-up: the final balance still runs.
            print(f'WARNING: make-up failed: {e!r}. Carrying on without it.')
            runs.append(scale.Run('makeup', note=f'failed: {e!r}'))
    if run['balance_swing']:
        try:
            runs += area.balance_swing(SimAuto, targets, epc=reports['swing'])
        except Exception as e:     # Last resort: balance_swing() keeps its own failed rounds; never lose the hour.
            print(f'WARNING: final balance failed: {e!r}. Saving the case as it is.')
            runs.append(scale.Run('balance', note=f'failed: {e!r}'))
    # Read back from PowerWorld, so the log shows what the case holds rather than what the scaling meant to do.
    final = area.read_case(SimAuto) if area_mode else {t: pw.get_df(SimAuto, t, f) for t, f in topo.TARGET_FIELDS.items()}
    final_loss = area.area_losses(SimAuto) if area_mode else None      # The areas sheet's Final Loss MW.
    swing_buses = prep.slack_buses(SimAuto)
    if not pw.solve(SimAuto):
        print('WARNING: the case does not solve at the end of scaling. Saved as it is.')
    # Save the case first: a log which can't be built or written (e.g. open in Excel) mustn't cost the solved hour.
    pw.save_case(SimAuto, out_fp, run['case_format'])
    summary = None
    try:
        log = build_log(runs, final, swing_buses, final_loss=final_loss)
        summary = summary_line(gv_fp.stem, log)
        print(summary)
        pw.write_excel(log_fp, log)
    except Exception as e:     # Never lose the hour over its log.
        print(f'WARNING: could not write {log_fp.name} (open in Excel?): {e!r}. The case was saved; '
              f'delete {out_fp.name} to rerun this hour with its log.')

    SimAuto.CloseCase()
    return summary

def check_inputs():
    """Stops with a clear message if step 01 didn't leave what this run's settings need."""
    if not pw_fp.exists():
        raise SystemExit(f'Step 02 needs {pw_fp}: run step 01 (prepare) first.')
    changes = prepared_changes()
    if changes and prep_stamp_fp.exists():
        raise SystemExit(f'{pw_fp.name} was made from other settings or inputs than now: ' + '; '.join(changes)
                         + '. Run step 01 (prepare): it redoes it.')
    if changes:     # Made before the record existed: it may be fine.
        print(f'WARNING: {changes[0]}: if settings.toml or the seed changed since it was made, run step 01 (prepare).')
    if run['use_pvqv'] and not pvqv_fp.exists():
        raise SystemExit(f'use_pvqv needs {pvqv_fp}: run step 01 with compute_pvqv = true.')
    if run['mode'] == 'topology' and not dummy_bus_fp.exists():
        raise SystemExit(f'Topology mode needs {dummy_bus_fp}: run step 01 in topology mode.')

def hours_to_scale() -> list[tuple[Path, Path]]:
    """(EPC, output PWB) for every hour EPC without an output PWB yet."""
    todo = []
    for gv_fp in sorted(gv_dir.glob('*.epc')):
        out_fp = out_dir / (gv_fp.stem + '.pwb')
        if out_fp.exists():
            print(f'Skipping {gv_fp.name}: {out_fp.name} already exists. Delete it to re-run.')
        else:
            todo.append((gv_fp, out_fp))
    return todo

def scale_hours(workers: int = 1):
    check_inputs()
    todo = hours_to_scale()
    if not todo:
        print('No hours to scale.')
        return
    out_dir.mkdir(exist_ok=True)
    workers = min(workers, len(todo))
    if workers > 1:
        return _scale_hours_parallel(todo, workers)
    SimAuto = pw.connect()
    for n, (gv_fp, out_fp) in enumerate(todo, 1):
        print(f'=== {gv_fp.name} ({n} of {len(todo)}) ===')
        # Keep going with the remaining hours if one fails.
        try:
            create_case(SimAuto, gv_fp, out_fp)
        except Exception:
            print(f'FAILED: {gv_fp.name}')
            traceback.print_exc()

# ------------------ Parallel: one process and PowerWorld instance per hour ------------------
def _init_worker(workers: int):
    """Runs in each worker process (which has re-imported this module, so the settings are loaded)."""
    # pretest_gens starts its own PowerWorld instances: share the CPUs out between the hours running at once.
    scale.MAX_WORKERS = max(1, (scale.MAX_WORKERS or os.cpu_count() or 1) // workers)

def _scale_one(gv_fp: Path, out_fp: Path, console_fp: Path) -> tuple[str | None, str | None]:
    """One hour in a worker, its console written to console_fp. Returns (summary line, error text)."""
    with open(console_fp, 'w', encoding='utf-8', buffering=1) as f, contextlib.redirect_stdout(f), contextlib.redirect_stderr(f):
        print(f'=== {gv_fp.name} ===')
        try:
            # Connected here, so a failed connect (e.g. no free SimAuto seat) costs this hour only.
            return create_case(pw.connect(), gv_fp, out_fp), None
        except Exception as e:
            traceback.print_exc()
            return None, repr(e)

def _worker(gv_fp: Path, out_fp: Path, console_fp: Path, workers: int, conn):
    """A worker process: one hour, its result sent back on conn."""
    _init_worker(workers)
    conn.send(_scale_one(gv_fp, out_fp, console_fp))
    conn.close()

_MAIN_STEP = re.compile(r'^Main scaling: \d+ -> (\d+) of (\d+)')
_MIDRUN = re.compile(r'\(main scaling at (\d+) of (\d+)\)')
_ANY_STEP = re.compile(r'^([^:]+?)(?: \(main scaling at [^)]*\))?: \d+ -> \d+ of \d+')
MAIN_SHARE = 95         # The progress bar counts an hour's main scaling as 95 of its 100; the runs after it, the rest.

def hour_progress(console_fp: Path) -> tuple[int, str]:
    """
    (0..MAIN_SHARE, what it is doing) for a running hour, from the end of its console: how far its main scaling has
    got, and the label of its latest scaling step (e.g. 'Main scaling', 'Balance 2', 'Final balance 1').
    """
    try:
        with open(console_fp, 'rb') as f:
            f.seek(0, os.SEEK_END)
            f.seek(max(0, f.tell() - 16384))
            lines = f.read().decode('utf-8', 'replace').splitlines()
    except OSError:
        return 0, 'starting'
    stage = next((m[1] for line in reversed(lines) if (m := _ANY_STEP.match(line))), 'planning')
    for line in reversed(lines):
        if line.startswith('Main scaling: done'):
            return MAIN_SHARE, stage
        if (m := _MAIN_STEP.match(line) or _MIDRUN.search(line)):
            return MAIN_SHARE * int(m[1]) // int(m[2]), stage
    return 0, stage

def _scale_hours_parallel(todo: list[tuple[Path, Path]], workers: int):
    """
    Each hour in a process of its own (spawn, as on Windows: it re-imports this module) with its own PowerWorld, at most
    `workers` at once. A process which dies (e.g. PowerWorld takes it down) costs only its own hour. A progress bar
    (tqdm) counts 100 per hour, moving with each running hour's main scaling, and shows what each is doing.
    """
    from tqdm import tqdm
    print(f'Scaling {len(todo)} hours, {workers} at a time. Each hour\'s console: {out_dir / "<hour>_Console.txt"}')
    ctx = mp.get_context('spawn')
    pending, running, n, failed = list(todo), {}, 0, 0
    seen: dict[str, int] = {}       # Hour -> its progress so far (never goes back on the bar).
    with tqdm(total=100 * len(todo), unit='%', dynamic_ncols=True, smoothing=0.05,
              bar_format='{percentage:3.0f}% |{bar}| {elapsed}<{remaining} {desc}') as bar:
        def redraw():
            hours = []
            for _, _, gv_fp, console_fp in running.values():
                progress, stage = hour_progress(console_fp)
                seen[gv_fp.stem] = max(seen.get(gv_fp.stem, 0), progress)
                hours.append(f'{gv_fp.stem} {stage} {progress}%')
            bar.update(100 * n + sum(seen[gv.stem] for _, _, gv, _ in running.values()) - bar.n)
            bar.set_description_str(f'{n}/{len(todo)} done' + (f', {failed} failed' if failed else '') + ' | ' + ' | '.join(hours))

        while pending or running:
            while pending and len(running) < workers:
                gv_fp, out_fp = pending.pop(0)
                console_fp = out_dir / (gv_fp.stem + '_Console.txt')
                recv, send = ctx.Pipe(duplex=False)
                p = ctx.Process(target=_worker, args=(gv_fp, out_fp, console_fp, workers, send))
                p.start()
                send.close()        # The child's end: recv then sees EOF if the child dies without sending.
                running[p.sentinel] = (p, recv, gv_fp, console_fp)
            for sentinel in multiprocessing.connection.wait(list(running), timeout=2.0):
                p, recv, gv_fp, _ = running.pop(sentinel)
                try:
                    summary, error = recv.recv()
                except EOFError:
                    summary, error = None, None
                p.join()
                recv.close()
                n += 1
                seen.pop(gv_fp.stem, None)
                if error is None and summary is None and p.exitcode != 0:
                    error = f'its process died (exit code {p.exitcode})'
                if error:
                    failed += 1
                    bar.write(f'[{n}/{len(todo)}] FAILED: {gv_fp.name}: {error} (see {gv_fp.stem}_Console.txt)')
                else:
                    bar.write(f'[{n}/{len(todo)}] {summary or gv_fp.stem + ": done (no log, see its console)"}')
            redraw()

# ================== Step 03: merge the hour logs ==================
EXCEL_MAX_ROWS = 1_048_576
MERGED_SHEETS = ['areas', 'exceptions', 'runs', 'shortfall']   # Fixed sheet names written by step 02 (topology mode: no areas, shortfall).
HOUR_COLUMNS = ['Hour', 'Status', 'Final Swing MW', 'Unmet MW', 'Not Reached MW', 'Make-up MW', 'Balance MW', 'Worst Error MW', 'Worst AreaNumber',
                'Worst Group', 'Exceptions Gen/Load', 'Exceptions STATCOM', 'Exceptions Branch']

def _hour_row(hour: str, sheets: dict[str, pd.DataFrame]) -> dict:
    """One `hours` row from one hour's sheets (missing or empty sheets give blanks, or 0 for counts)."""
    row: dict[str, object] = dict.fromkeys(HOUR_COLUMNS, np.nan)
    row['Hour'] = hour
    runs = sheets.get('runs')
    # Step 02 writes the plan's log before scaling: without a `main` run the hour failed or never finished, and the
    # columns below are the plan's (or blank), not a result.
    if runs is None or 'Run' not in runs or not (runs['Run'] == 'main').any():
        row['Status'] = 'incomplete'
    else:       # Scaled, but a later run (make-up, dist gen, a balance) failed or stopped: the swing may be off its goal.
        notes = runs['Note'].astype(str) if 'Note' in runs else pd.Series(dtype=str)
        row['Status'] = 'check' if notes.str.contains(r'(?:^|; )(?:failed|stopped)').any() else 'ok'
    # Near 0 with a giant swing; with the case's own slack, its end MW (its goal is its EPC MW, or within its limits).
    row['Final Swing MW'] = runs['Swing MW After'].iloc[-1] if runs is not None and len(runs) else np.nan
    areas = sheets.get('areas')
    if areas is not None and len(areas):
        for col in ('Unmet MW', 'Not Reached MW', 'Make-up MW'):
            # Unmet: what groups' own units couldn't take; Make-up covers some of it (and Not Reached) within each area.
            row[col] = areas[col].sum() if col in areas else np.nan      # Make-up MW: not in older logs.
        row['Balance MW'] = areas.loc[areas['Group'] != 'Load', 'Balance MW'].sum()     # Balancing moves gens only.
        err = areas['Error MW'].abs()
        if err.notna().any():
            worst = areas.iloc[int(np.nanargmax(err.to_numpy()))]
            row['Worst Error MW'] = worst['Error MW']
            row['Worst AreaNumber'] = worst['AreaNumber']
            row['Worst Group'] = worst['Group']
    exc = sheets.get('exceptions')
    if exc is not None and len(exc):
        row['Exceptions Gen/Load'] = int(exc['Table'].isin(['Gen', 'Load']).sum())
        row['Exceptions STATCOM'] = int((exc['Table'] == 'STATCOM').sum())
        row['Exceptions Branch'] = int((exc['Table'] == 'Branch').sum())
    else:
        row['Exceptions Gen/Load'] = row['Exceptions STATCOM'] = row['Exceptions Branch'] = 0
    return row

def merge(out_dir: Path = out_dir) -> Path:
    """
    Merges every <hour>_Log.xlsx in out_dir. Returns the workbook path ('03 Merge Reports.xlsx', beside out_dir).
    Sheets: `hours` (one row per hour; Status 'incomplete' = no `main` run in its log: the hour's areas rows are still
    merged, but show the plan; 'check' = scaled, but a later run failed or stopped, so the swing may be off its goal),
    then each of MERGED_SHEETS with an Hour column first. A merged sheet over Excel's row limit is written as a CSV
    beside the workbook instead, and printed.
    """
    files = sorted(f for f in Path(out_dir).glob('*_Log.xlsx') if not f.name.startswith('~$'))  # Not Excel's lock files.
    if not files:
        raise SystemExit(f'No *_Log.xlsx files found in {out_dir}')
    target_fp = Path(out_dir).parent / '03 Merge Reports.xlsx'

    frames: dict[str, list[pd.DataFrame]] = {name: [] for name in MERGED_SHEETS}
    hour_rows = []
    for file in files:
        hour = file.stem[:-len('_Log')]
        sheets = pd.read_excel(file, sheet_name=None)
        hour_rows.append(_hour_row(hour, sheets))
        for name in MERGED_SHEETS:
            df = sheets.get(name)
            if df is None or df.empty:
                continue
            df.insert(0, 'Hour', hour)
            frames[name].append(df)

    out: dict[str, pd.DataFrame] = {'hours': pd.DataFrame(hour_rows, columns=HOUR_COLUMNS)}
    for name in MERGED_SHEETS:
        if not frames[name]:
            out[name] = pd.DataFrame(columns=['Hour'])
            continue
        merged = pd.concat(frames[name], ignore_index=True)
        if len(merged) + 1 > EXCEL_MAX_ROWS:
            csv_fp = Path(out_dir).parent / f'03 Merge {name}.csv'
            merged.to_csv(csv_fp, index=False)
            print(f"'{name}' has {len(merged)} rows, over Excel's limit. Written to '{csv_fp}' instead.")
        else:
            out[name] = merged
    pw.write_excel(target_fp, out)
    print(f"Concatenation complete. Data saved to '{target_fp}'.")
    return target_fp

# ================== The run ==================
def main(argv=None):
    parser = argparse.ArgumentParser(description=(__doc__ or '').strip().splitlines()[0])
    parser.add_argument('steps', nargs='*', metavar='STEP',
                        help=f'Steps to run, from {", ".join(STEPS)}. Default: [run] steps in settings.toml.')
    args = parser.parse_args(argv)
    steps = args.steps or run['steps']
    if set(steps) - set(STEPS):
        parser.error(f'unknown steps {sorted(set(steps) - set(STEPS))}: choose from {", ".join(STEPS)}')
    workers = run['workers']
    print(f'Mode: {run["mode"]}; steps: {", ".join(s for s in STEPS if s in steps)}')

    if 'prepare' in steps:
        print('=== Step 01: prepare ===')
        prepare_if_needed()
    if 'scale' in steps:
        print(f'=== Step 02: scale ({workers} at a time) ===')
        scale_hours(workers)
    if 'merge' in steps:
        print('=== Step 03: merge ===')
        merge(out_dir)
    print('done')

if __name__ == '__main__':
    main()
