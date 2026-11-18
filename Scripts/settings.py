"""
Reads settings.toml (repo root). No dependencies: tomllib is in Python 3.11+.

[run] holds the switches pcm_to_pf.py reads, returned by load(). Every other section sets the constants of the module
it names, by upper-cased key: [scale] exclude_vmin = 0.88 sets Scripts/scale.py's EXCLUDE_VMIN. The modules hold the
defaults, so a key left out of the file keeps its default, and an unknown key (one not in SETTINGS) is an error.
"""
import tomllib
from pathlib import Path
from . import area, prep, pw, scale, topo

DEFAULT_FP = Path(__file__).resolve().parents[1] / 'settings.toml'

MODULES = {'pw': pw, 'prep': prep, 'topology': topo, 'area': area, 'scale': scale}

RUN_DEFAULTS = {
    'mode': 'area',                         # 'area' or 'topology'
    'case_format': 'PWB23',
    'create_giant_swing': True,
    'swing_bus': 'slack',                   # 'slack', 'max_fault_duty', or a bus number
    'adjust_seed_shunts': True,
    'gen_terminal_voltage_control': True,
    'compute_pvqv': True,
    'pretest_gens': False,
    'use_pvqv': False,
    'balance_swing': True,
    'steps': ['prepare', 'scale', 'merge'],  # pcm_to_pf.py's steps when run without arguments.
    'workers': 1,                           # Hours scaled at once, each in its own PowerWorld instance.
}
MODES = ('area', 'topology')
STEPS = ('prepare', 'scale', 'merge')
# The keys each section accepts: those in the shipped settings.toml (max_workers is commented out there). Other upper-case
# names in the modules (RESOLUTION, STATCOM_ID, LOAD_P, STATUS, ...) are code, not settings.
SETTINGS = {
    'pw': {'mva_mismatch_threshold'},
    'prep': {'swing_mw_max', 'swing_mw_min', 'shunt_vlow', 'shunt_vhigh', 'shunt_max_iterations'},
    'topology': {'tap_threshold', 'unsupported_gv_branch_types'},
    'area': {'commit', 'commit_max_cf', 'capacity_types', 'off_at_zero_types', 'manual_unit_types', 'scale_nonconforming',
             'loss_exponent', 'makeup', 'balance_tol_mw', 'balance_max_rounds', 'balance_trigger_mw',
             'balance_at_swing_limits', 'max_unknown_share', 'shortfall_report_mw', 'gen_types', 'unit_gentype', 'fuel_gentype'},
    'scale': {'add_statcoms', 'drop_collapsed', 'exclude_vmin', 'exclude_vmax', 'statcom_vpu', 'statcom_vnom_min',
              'collapse_vpu', 'pvqv_dv_limit', 'initial_step', 'max_step', 'step_solves', 'final_max_rounds', 'max_workers'},
}
# Code letter sets / maps: compared against upper-cased codes.
CODE_KEYS = {(area, 'UNIT_GENTYPE'), (area, 'FUEL_GENTYPE'), (area, 'MANUAL_UNIT_TYPES')}
FUEL = 'fuel'   # In [area.unit_gentype]: decide by the fuel type.

def _check_type(fp, what: str, default, value, allow_int: bool = False):
    """The value must be of the default's kind: bool, number, text, list (for lists and sets) or table."""
    def kind(v):
        if isinstance(v, bool):
            return 'true/false'
        if isinstance(v, (int, float)):
            return 'number'
        if isinstance(v, str):
            return 'text'
        if isinstance(v, (list, set, tuple)):
            return 'list'
        if isinstance(v, dict):
            return 'table'
        return None     # e.g. None: anything goes.
    want, got = kind(default), kind(value)
    if want is not None and got != want and not (allow_int and got == 'number'):
        raise ValueError(f'{fp}: {what} must be {want}, not {value!r}')

def _need(fp, ok: bool, what: str):
    if not ok:
        raise ValueError(f'{fp}: {what}')

def _is_int(v) -> bool:
    return isinstance(v, int) and not isinstance(v, bool)

def _check_ranges(fp, get):
    """Rejects values which hang the engine or silently disable it. `get(module, NAME)`: the value the file would set."""
    for name in ('INITIAL_STEP', 'MAX_STEP', 'STEP_SOLVES', 'FINAL_MAX_ROUNDS'):
        _need(fp, _is_int(get(scale, name)), f'[scale] {name.lower()} must be a whole number')
    _need(fp, _is_int(get(area, 'BALANCE_MAX_ROUNDS')), '[area] balance_max_rounds must be a whole number')
    workers = get(scale, 'MAX_WORKERS')
    _need(fp, workers is None or (_is_int(workers) and workers >= 1), f'[scale] max_workers must be a whole number of 1 or more, not {workers!r}')
    _need(fp, 1 <= get(scale, 'INITIAL_STEP') <= get(scale, 'MAX_STEP') <= scale.RESOLUTION,
          f'[scale] needs 1 <= initial_step <= max_step <= {scale.RESOLUTION}, not {get(scale, "INITIAL_STEP")}, {get(scale, "MAX_STEP")}')
    _need(fp, get(scale, 'STEP_SOLVES') >= 1, f'[scale] step_solves must be 1 or more, not {get(scale, "STEP_SOLVES")}')
    _need(fp, get(scale, 'FINAL_MAX_ROUNDS') >= 1, f'[scale] final_max_rounds must be 1 or more, not {get(scale, "FINAL_MAX_ROUNDS")}')
    _need(fp, get(scale, 'EXCLUDE_VMIN') < get(scale, 'EXCLUDE_VMAX'),
          f'[scale] exclude_vmin ({get(scale, "EXCLUDE_VMIN")}) must be below exclude_vmax ({get(scale, "EXCLUDE_VMAX")})')
    _need(fp, get(pw, 'MVA_MISMATCH_THRESHOLD') > 0, f'[pw] mva_mismatch_threshold must be above 0, not {get(pw, "MVA_MISMATCH_THRESHOLD")}')
    _need(fp, 0 < get(area, 'COMMIT_MAX_CF') <= 1, f'[area] commit_max_cf must be above 0 and at most 1, not {get(area, "COMMIT_MAX_CF")}')
    _need(fp, get(area, 'BALANCE_MAX_ROUNDS') >= 0, f'[area] balance_max_rounds must be 0 or more, not {get(area, "BALANCE_MAX_ROUNDS")}')
    _need(fp, get(area, 'BALANCE_TRIGGER_MW') >= 0, f'[area] balance_trigger_mw must be 0 or more, not {get(area, "BALANCE_TRIGGER_MW")}')
    _need(fp, get(area, 'BALANCE_TOL_MW') > 0, f'[area] balance_tol_mw must be above 0, not {get(area, "BALANCE_TOL_MW")}')
    _need(fp, get(prep, 'SWING_MW_MIN') < get(prep, 'SWING_MW_MAX'),
          f'[prep] swing_mw_min ({get(prep, "SWING_MW_MIN")}) must be below swing_mw_max ({get(prep, "SWING_MW_MAX")})')
    _need(fp, get(area, 'SHORTFALL_REPORT_MW') >= 0, f'[area] shortfall_report_mw must be 0 or more, not {get(area, "SHORTFALL_REPORT_MW")}')
    _need(fp, get(area, 'LOSS_EXPONENT') >= 0, f'[area] loss_exponent must be 0 or more, not {get(area, "LOSS_EXPONENT")}')
    gen_types = set(get(area, 'GEN_TYPES'))
    bad = {v for n in ('UNIT_GENTYPE', 'FUEL_GENTYPE') for v in get(area, n).values() if v is not None and v not in gen_types}
    bad |= (get(area, 'CAPACITY_TYPES') | get(area, 'OFF_AT_ZERO_TYPES')) - gen_types
    _need(fp, not bad, f'GenTypes {sorted(bad)} are not in [area] gen_types {get(area, "GEN_TYPES")}')

def load(fp: Path = DEFAULT_FP) -> dict:
    """
    Applies the file's module settings, and returns the [run] settings (defaults filled in). Every value is checked
    before any is applied: a rejected file changes nothing.
    """
    data = tomllib.loads(Path(fp).read_text(encoding='utf-8'))
    run = dict(RUN_DEFAULTS)
    new = {}        # (module, NAME) -> the value to set.
    for section, values in data.items():
        if section == 'run':
            unknown = set(values) - set(RUN_DEFAULTS)
            if unknown:
                raise ValueError(f'{fp}: unknown [run] settings: {sorted(unknown)}')
            for key, value in values.items():
                _check_type(fp, f'[run] {key}', RUN_DEFAULTS[key], value, allow_int=key == 'swing_bus')
            run.update(values)
            continue
        if section not in MODULES:
            raise ValueError(f'{fp}: unknown section [{section}]. Sections: run, {", ".join(MODULES)}')
        module = MODULES[section]
        for key, value in values.items():
            if key not in SETTINGS[section]:
                raise ValueError(f'{fp}: unknown setting [{section}] {key}')
            name = key.upper()
            current = getattr(module, name)
            _check_type(fp, f'[{section}] {key}', current, value)
            if (module, name) in CODE_KEYS:
                value = {k.upper(): v for k, v in value.items()} if isinstance(value, dict) else [v.upper() for v in value]
            if isinstance(current, dict):
                # Tables add to / override the default map, rather than replace it.
                value = {**current, **{k: (None if v == FUEL else v) for k, v in value.items()}}
            elif isinstance(current, set):
                value = set(value)
            new[(module, name)] = value

    if run['mode'] not in MODES:
        raise ValueError(f'{fp}: [run] mode must be one of {MODES}, not {run["mode"]!r}')
    if not (run['swing_bus'] in ('slack', 'max_fault_duty') or type(run['swing_bus']) is int):
        raise ValueError(f"{fp}: [run] swing_bus must be 'slack', 'max_fault_duty' or a bus number")
    if not run['steps'] or set(run['steps']) - set(STEPS):
        raise ValueError(f'{fp}: [run] steps must be from {STEPS}, not {run["steps"]!r}')
    if not (_is_int(run['workers']) and run['workers'] >= 1):
        raise ValueError(f'{fp}: [run] workers must be a whole number of 1 or more, not {run["workers"]!r}')
    _check_ranges(fp, lambda module, name: new.get((module, name), getattr(module, name)))
    for (module, name), value in new.items():
        setattr(module, name, value)
    return run
