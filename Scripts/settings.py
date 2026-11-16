"""
Reads settings.toml (repo root). No dependencies: tomllib is in Python 3.11+.

[run] holds the switches the step scripts read, returned by load(). Every other section sets the constants of the module
it names, by upper-cased key: [scale] exclude_vmin = 0.88 sets Scripts/scale.py's EXCLUDE_VMIN. The modules hold the
defaults, so a key left out of the file keeps its default, and an unknown key is an error.
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
}
MODES = ('area', 'topology')
FUEL = 'fuel'   # In [area.unit_gentype]: decide by the fuel type.

def load(fp: Path = DEFAULT_FP) -> dict:
    """Applies the file's module settings, and returns the [run] settings (defaults filled in)."""
    data = tomllib.loads(Path(fp).read_text(encoding='utf-8'))
    run = dict(RUN_DEFAULTS)
    for section, values in data.items():
        if section == 'run':
            unknown = set(values) - set(RUN_DEFAULTS)
            if unknown:
                raise ValueError(f'{fp}: unknown [run] settings: {sorted(unknown)}')
            run.update(values)
            continue
        if section not in MODULES:
            raise ValueError(f'{fp}: unknown section [{section}]. Sections: run, {", ".join(MODULES)}')
        module = MODULES[section]
        for key, value in values.items():
            name = key.upper()
            if not hasattr(module, name):
                raise ValueError(f'{fp}: unknown setting [{section}] {key}')
            current = getattr(module, name)
            if isinstance(current, dict):
                # Tables add to / override the default map, rather than replace it.
                value = {**current, **{k: (None if v == FUEL else v) for k, v in value.items()}}
            elif isinstance(current, set):
                value = set(value)
            setattr(module, name, value)

    if run['mode'] not in MODES:
        raise ValueError(f'{fp}: [run] mode must be one of {MODES}, not {run["mode"]!r}')
    if not (run['swing_bus'] in ('slack', 'max_fault_duty') or isinstance(run['swing_bus'], int)):
        raise ValueError(f"{fp}: [run] swing_bus must be 'slack', 'max_fault_duty' or a bus number")
    gen_types = set(area.GEN_TYPES)
    bad = {v for m in (area.UNIT_GENTYPE, area.FUEL_GENTYPE) for v in m.values() if v is not None and v not in gen_types}
    bad |= (area.ALL_ON_TYPES | area.OFF_AT_ZERO_TYPES) - gen_types
    if bad:
        raise ValueError(f'{fp}: GenTypes {sorted(bad)} are not in [area] gen_types {area.GEN_TYPES}')
    return run
