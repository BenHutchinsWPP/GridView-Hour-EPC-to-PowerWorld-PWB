"""
Everything which talks to PowerWorld through SimAuto, plus the Excel report writer.
Every SimAuto error raises SimAutoError.
"""
from collections.abc import Iterator, Sequence
from pathlib import Path
from contextlib import contextmanager
from types import SimpleNamespace
from typing import Any, Protocol, cast
import pandas as pd
import openpyxl.utils

MVA_MISMATCH_THRESHOLD = 1.0    # Max bus MVA mismatch for a solve to count as converged.

class SimAutoApi(Protocol):
    """
    The part of PowerWorld's SimAuto COM object this project uses (connect() returns the real one; the tests pass
    fakes). Each call returns a tuple: [0] an error message ('' if none), then any data (chk() unpacks it).
    Positional-only, so a fake's parameter names don't have to match.
    """
    CreateIfNotFound: bool

    def OpenCase(self, fp: str, /) -> tuple[Any, ...]: ...
    def SaveCase(self, fp: str, case_format: str, overwrite: bool, /) -> tuple[Any, ...]: ...
    def CloseCase(self) -> tuple[Any, ...]: ...
    def RunScriptCommand(self, command: str, /) -> tuple[Any, ...]: ...
    def GetParametersMultipleElementRect(self, table: str, fields: list[str], filter_group: str, /) -> tuple[Any, ...]: ...
    def ChangeParametersMultipleElementRect(self, table: str, fields: list[str], rows: list[list[Any]], /) -> tuple[Any, ...]: ...
    def SaveState(self) -> tuple[Any, ...]: ...
    def LoadState(self) -> tuple[Any, ...]: ...

class SimAutoError(Exception):
    log: dict | None = None     # scale.iterate_to_gen_load_targets() hands its run's log on with it.

def chk(SimAutoOutput: Sequence[Any], what: str) -> Any:
    """
    Raises SimAutoError on any error passed back from SimAuto, otherwise returns the data.

    SimAuto return object format:
    [0] = Error message, if any
    [1+] = Return data
    """
    if SimAutoOutput[0] != '':
        raise SimAutoError(f'{what}: {SimAutoOutput[0]}')

    if len(SimAutoOutput) == 1:
        return None
    elif len(SimAutoOutput) == 2:
        return SimAutoOutput[1]
    else:
        return SimAutoOutput[1:]

def connect() -> SimAutoApi:
    """Starts PowerWorld through SimAuto (Windows only: pywin32 is imported here, so nothing else needs it to import)."""
    import win32com.client
    return cast(SimAutoApi, win32com.client.Dispatch("pwrworld.SimulatorAuto"))

def run(SimAuto: SimAutoApi, command: str) -> Any:
    """Runs a script command."""
    return chk(SimAuto.RunScriptCommand(command), command)

def get_df(SimAuto: SimAutoApi, table: str, fields: dict[str,type], filter_group: str = '') -> pd.DataFrame:
    """Reads `fields` of every object in `table` (optionally filtered), converted to the given types. Strings are trimmed."""
    rows = chk(SimAuto.GetParametersMultipleElementRect(table, list(fields), filter_group), f'Get {table}')
    df = pd.DataFrame(data=rows, columns=list(fields))
    for field, typ in fields.items():
        if typ in [int, float]:
            # Handle numeric types with coerce for empty strings.
            df[field] = pd.to_numeric(df[field], errors='coerce')
            if typ is float:
                # Whole-number values would otherwise come back as int64, which float values can't later be written into.
                df[field] = df[field].astype(float)
        else:
            df[field] = df[field].astype(str).str.strip()      # Every non-numeric field is text.
    return df

def set_df(SimAuto: SimAutoApi, table: str, df: pd.DataFrame, fields: list[str]) -> None:
    """Writes only `fields` of df. A named index (e.g. ObjectID) is available as a field. Missing values are sent as ''."""
    if len(df) == 0:
        return
    rows = [['' if pd.isna(v) else str(v) for v in row] for row in df.reset_index()[fields].itertuples(index=False)]
    chk(SimAuto.ChangeParametersMultipleElementRect(table, fields, rows), f'Set {table}')

@contextmanager
def trial(SimAuto: SimAutoApi, solves: int = 1, keep_changes: bool = True) -> Iterator[SimpleNamespace]:
    """
    Checkpoints the case, runs the body, then solves.
    Rolls back to the checkpoint if the case does not solve (or if keep_changes=False).
    Uses SimAuto's one in-memory checkpoint (SaveState/LoadState), which works in Edit and Run mode alike.
    The named-state script commands (StoreState etc.) don't: they need Run mode, and switching
    from Edit to Run destroys them. With one checkpoint, trials can't nest.

        with trial(SimAuto) as t:
            set_df(...)
        if not t.ok: ...
    """
    global _in_trial
    if _in_trial:
        raise RuntimeError('pw.trial cannot nest: SimAuto has only one checkpoint.')
    _in_trial = True
    try:
        chk(SimAuto.SaveState(), 'SaveState')
        t = SimpleNamespace(ok=False)
        try:
            yield t
            t.ok = all(solve(SimAuto) for _ in range(solves))
        finally:
            if not (t.ok and keep_changes):
                chk(SimAuto.LoadState(), 'LoadState')
    finally:
        _in_trial = False

_in_trial = False

@contextmanager
def creating(SimAuto: SimAutoApi) -> Iterator[None]:
    """Allows set_df to create objects which don't exist yet, only within this block."""
    run(SimAuto, 'EnterMode(EDIT);')
    SimAuto.CreateIfNotFound = True
    try:
        yield
    finally:
        SimAuto.CreateIfNotFound = False

def set_df_bisect(SimAuto: SimAutoApi, table: str, df: pd.DataFrame, fields: list[str], solves: int = 1, limit: int | None = None) -> list:
    """
    Attempts to apply all rows of df in one bulk edit.
    If it fails, rolls back and bisects df into two halves to test recursively until reaching 1 row.
    Each attempt must pass `solves` solves.
    Returns the index labels of rows which could not be applied without divergence.
    limit: stop bisecting once this many single rows have failed. Rows not yet tried are returned as not applied too.
    """
    failed: list = []

    def apply(part: pd.DataFrame) -> None:
        if len(part) == 0:
            return
        if limit is not None and len(failed) >= limit:
            failed.extend(part.index)       # Not tried.
            return
        with trial(SimAuto, solves=solves) as t:
            set_df(SimAuto, table, part, fields)
        if t.ok:
            return
        if len(part) == 1:
            failed.append(part.index[0])
            return
        mid = len(part) // 2
        apply(part.iloc[:mid])
        apply(part.iloc[mid:])

    apply(df)
    return failed

print_opened = True    # Step 02 turns this off: it opens the seed three times an hour.
print_solve_errors = True   # Step 02 turns this off: its step lines say which steps diverged.

def open_case(SimAuto: SimAutoApi, fp: str | Path) -> None:
    if not Path(fp).exists():
        raise FileNotFoundError(fp)
    chk(SimAuto.OpenCase(str(fp)), f'OpenCase {fp}')
    if print_opened:
        print(f'Opened: {str(fp)}')

def save_case(SimAuto: SimAutoApi, fp: str | Path, case_format: str = 'PWB22') -> None:
    if not Path(fp).parent.exists():
        raise FileNotFoundError(Path(fp).parent)
    chk(SimAuto.SaveCase(str(fp), case_format, True), f'SaveCase {fp}')
    print(f'Saved: {str(fp)}')

def solve(SimAuto: SimAutoApi) -> bool:
    run(SimAuto, 'EnterMode(RUN);')
    result = SimAuto.RunScriptCommand('SolvePowerFlow(RECTNEWT);')

    # Error string. Return early with False if it didn't solve.
    if result[0] != '':
        if print_solve_errors:
            print(result[0])
        return False
    run(SimAuto, 'EnterMode(EDIT);')

    # A "solved" case with large mismatches doesn't count.
    df = get_df(SimAuto, 'Bus', {'Busnum':int, 'MismatchP':float, 'MismatchQ':float})
    max_mismatch = ((df['MismatchP']**2.0 + df['MismatchQ']**2.0)**0.5).max()
    return bool(max_mismatch < MVA_MISMATCH_THRESHOLD)

def write_excel(fp: Path, sheets: dict[str,pd.DataFrame]) -> None:
    """Writes one sheet per dataframe. A named index (e.g. ObjectID) is written as a column."""
    with pd.ExcelWriter(fp, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=df.index.name is not None)
        for sheet in writer.sheets.values():
            # Auto-fit columns, freeze and filter the top row.
            for col in sheet.columns:
                width = max((len(str(cell.value)) for cell in col if cell.value is not None), default=0)
                sheet.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = width + 2
            sheet.freeze_panes = 'A2'
            sheet.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(sheet.max_column)}1"
