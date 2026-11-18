"""Shared fakes, table builders and settings helpers for the tests (not tests themselves)."""
import sys
import importlib.util
from typing import Any
from contextlib import contextmanager
from pathlib import Path
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from Scripts import pw, scale, settings

class FakeSimAuto:
    """
    In-memory stand-in for SimAuto. Records written rows by (table, first field value).
    SolvePowerFlow fails while any row with a key in `bad` is applied.
    Reports two healthy buses (1 and 2) at 1.0 PU, and no shunts or branches.
    """
    def __init__(self, bad=()):
        self.bad = set(bad)
        self.applied: dict[tuple[str, str], dict] = {}
        self.checkpoint: dict[tuple[str, str], dict] | None = None
        self.CreateIfNotFound = False

    def OpenCase(self, fp):     # Tests which open cases use test_steps.CaseSimAuto.
        raise NotImplementedError('FakeSimAuto has no cases to open')

    def SaveCase(self, fp, case_format, overwrite):
        raise NotImplementedError('FakeSimAuto has no case to save')

    def CloseCase(self):
        return ('',)

    def SaveState(self):
        self.checkpoint = dict(self.applied)
        return ('',)

    def LoadState(self):
        assert self.checkpoint is not None
        self.applied = dict(self.checkpoint)
        return ('',)

    def RunScriptCommand(self, command) -> tuple[str]:
        if command.startswith('SolvePowerFlow') and self.diverged():
            return ('Diverged',)
        return ('',)

    def diverged(self):
        return bool(self.bad & {k for _, k in self.applied})

    def ChangeParametersMultipleElementRect(self, table, fields, rows):
        for row in rows:
            self.applied[(table, row[0])] = dict(zip(fields, row))
        return ('',)

    def GetParametersMultipleElementRect(self, table, fields, filter_group):
        if table != 'Bus':
            return ('', [])
        values = {'Vpu': '1.0', 'IslandNumber': '1', 'BusIsStarBus:1': 'NO', 'BusNomVolt': '230'}
        return ('', [[str(n) if f in ('Busnum', 'Number') else values.get(f, '0') for f in fields] for n in (1, 2)])

LOAD_P_ALL = scale.SCALED['Load']

def targets_df(**cols):
    df = pd.DataFrame(cols).set_index('ObjectID')
    df['ExclusionReason'] = ''
    return df

def cell(frame: pd.DataFrame, *key) -> Any:
    """frame.loc[key] typed as Any (pandas-stubs types a cell as a scalar union, and a row as Series or DataFrame)."""
    return frame.loc[key]

def results(log):
    return log['step_df']['Result'].tolist()

def loads(rows):
    """rows: (ObjectID, AreaNumber, Status, SMW, IMW, ZMW, SMvar)."""
    df = pd.DataFrame(rows, columns=['ObjectID', 'AreaNumber', 'Status', 'SMW', 'IMW', 'ZMW', 'SMvar'])
    return df.assign(BusNum=1, ID='1', AGC='YES', IMvar=0.0, ZMvar=0.0, DistStatus='Closed', DistMWInput=5.0, DistMvarInput=1.0)

def gens(rows):
    """rows: (ObjectID, AreaNumber, ID, Status, UnitTypeCode, MWSetPoint, MWMax, MWMin)."""
    df = pd.DataFrame(rows, columns=['ObjectID', 'AreaNumber', 'ID', 'Status', 'UnitTypeCode', 'MWSetPoint', 'MWMax', 'MWMin'])
    return df.assign(BusNum=1, FuelTypeCode='')

class SwingSimAuto(FakeSimAuto):
    """
    A case with gens & loads on their own buses, and the giant swing (ID xx, bus 99) carrying load minus generation.
    `slack`: bus numbers which report Slack = YES. Gens on them carry the imbalance too (so they are the swing when there is no xx).
    Written rows (keyed by bus) overlay the starting tables; StoreState/RestoreState work on them as in FakeSimAuto.
    """
    islands: dict      # Optional (read with getattr): Bus -> IslandNumber text.
    losses: dict       # Optional: Area -> loss MW.
    count_dist: bool   # Optional: dist gen counted in the swing.

    def __init__(self, gen, load, slack=()):
        super().__init__()
        self.slack = set(slack)
        self.base = {'Gen': gen.set_index('BusNum', drop=False), 'Load': load.set_index('BusNum', drop=False)}

    def table(self, name):
        df = self.base[name].copy()
        for (t, key), row in self.applied.items():
            if t == name:
                for f, v in row.items():
                    if f in df.columns:
                        df.loc[int(key), f] = type(df[f].iloc[0])(v) if f not in ('Status', 'DistStatus', 'ID') else v
        if name == 'Load':      # Solution columns at 1.0 PU: what the load draws, and what dist gen gives back.
            on = df['Status'] == 'Closed'
            df['MW'] = (df['SMW'] + df['IMW'] + df['ZMW']).where(on, 0.0)
            df['Mvar'] = (df['SMvar'] + df['IMvar'] + df['ZMvar']).where(on, 0.0)
            dist = on & (df['DistStatus'] == 'Closed')
            df['DistMW'], df['DistMvar'] = df['DistMWInput'].where(dist, 0.0), df['DistMvarInput'].where(dist, 0.0)
            df['NetMW'], df['NetMvar'] = df['MW'] - df['DistMW'], df['Mvar'] - df['DistMvar']
        if name == 'Gen':
            on = df['Status'] == 'Closed'
            ld = self.table('Load')
            p = ld.loc[ld['Status'] == 'Closed', ['SMW', 'IMW', 'ZMW']].sum().sum()
            if getattr(self, 'count_dist', False):     # Opt-in: dist gen injects while its load and itself are Closed.
                p -= ld.loc[(ld['Status'] == 'Closed') & (ld['DistStatus'] == 'Closed'), 'DistMWInput'].sum()
            df['MW'] = df['MWSetPoint']
            sw = (df['ID'] == 'xx') | df['BusNum'].isin(self.slack)
            df.loc[sw, 'MW'] = p - df.loc[on & ~sw, 'MWSetPoint'].sum()
        return df

    def GetParametersMultipleElementRect(self, table, fields, filter_group):
        if table in self.base:
            return ('', [[str(v) for v in row] for row in self.table(table)[fields].itertuples(index=False)])
        if table == 'Area':
            return ('', [[str(n), str(mw)] for n, mw in getattr(self, 'losses', {}).items()])
        if table == 'Bus':
            buses = sorted(set(self.base['Gen']['BusNum']) | set(self.base['Load']['BusNum']))
            values = {'Vpu': '1.0', 'IslandNumber': '1', 'BusIsStarBus:1': 'NO', 'BusNomVolt': '230'}
            islands = getattr(self, 'islands', {})      # Bus -> IslandNumber text ('' = blank); 1 if not given.
            return ('', [[str(n) if f in ('Busnum', 'Number') else islands.get(n, '1') if f == 'IslandNumber'
                          else 'YES' if f == 'Slack' and n in self.slack else values.get(f, 'NO' if f == 'Slack' else '0')
                          for f in fields] for n in buses])
        return ('', [])

def swing_case():
    gen = gens([('xx', 9, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '1', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[99, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    return gen, load

@contextmanager
def saved_settings():
    """Module settings (upper-case globals, and pw's print flags) as they are now, put back afterwards."""
    saved = module_settings()
    prints = pw.print_opened, pw.print_solve_errors
    try:
        yield
    finally:
        restore_settings(saved)
        pw.print_opened, pw.print_solve_errors = prints

def module_settings():
    return {(m, k): v for m in settings.MODULES.values() for k, v in vars(m).items() if k.isupper()}

def restore_settings(saved):
    for (m, k), v in saved.items():
        setattr(m, k, v)

def load_step_02(folder: Path, mode: str) -> Any:
    """A fresh pcm_to_pf.py, pointed at `folder`. SimAuto isn't connected: no win32com needed."""
    spec = importlib.util.spec_from_file_location('step_02', Path(__file__).resolve().parents[1] / 'pcm_to_pf.py')
    assert spec is not None and spec.loader is not None
    step: Any = importlib.util.module_from_spec(spec)
    saved = module_settings()
    spec.loader.exec_module(step)           # Loads settings.toml: the owner's choices must not reach the tests.
    restore_settings(saved)
    step.run = dict(settings.RUN_DEFAULTS, mode=mode)
    step.gv_dir, step.out_dir = folder / 'HourEPCs', folder / 'Output'
    step.pw_fp, step.pvqv_fp, step.dummy_bus_fp = folder / 'Prepared.pwb', folder / 'pvqv.csv', folder / 'DummyBus.aux'
    step.seed_dir, step.prep_stamp_fp = folder / 'Seed', folder / 'Prepared_Settings.json'
    for d in (step.gv_dir, step.out_dir):
        d.mkdir()
    for fp in (step.pw_fp, step.gv_dir / 'h1.epc'):
        fp.write_text('')
    return step
