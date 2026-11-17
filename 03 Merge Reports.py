from pathlib import Path
import numpy as np
import pandas as pd
from Scripts import pw

EXCEL_MAX_ROWS = 1_048_576
MERGED_SHEETS = ['areas', 'exceptions', 'runs']     # Fixed sheet names written by step 02 (area mode: all three; topology: no areas).
HOUR_COLUMNS = ['Hour', 'Final Swing MW', 'Unmet MW', 'Not Reached MW', 'Make-up MW', 'Balance Gen MW', 'Balance Load MW', 'Worst Error MW', 'Worst AreaNumber',
                'Worst Group', 'Exceptions Gen/Load', 'Exceptions STATCOM', 'Exceptions Branch']

def _hour_row(hour: str, sheets: dict[str, pd.DataFrame]) -> dict:
    """One `hours` row from one hour's sheets (missing or empty sheets give blanks, or 0 for counts)."""
    row = dict.fromkeys(HOUR_COLUMNS, np.nan)
    row['Hour'] = hour
    runs = sheets.get('runs')
    # Near 0 with a giant swing; with the case's own slack, its end MW (its goal is its EPC MW, or within its limits).
    row['Final Swing MW'] = runs['Swing MW After'].iloc[-1] if runs is not None and len(runs) else np.nan
    areas = sheets.get('areas')
    if areas is not None and len(areas):
        for col in ('Unmet MW', 'Not Reached MW', 'Make-up MW'):
            # Unmet: what groups' own units couldn't take; Make-up covers some of it (and Not Reached) within each area.
            row[col] = areas[col].sum() if col in areas else np.nan      # Make-up MW: not in older logs.
        # Gen and load balancing apart: summed together, MW put onto gen and taken off load would cancel.
        is_load = areas['Group'] == 'Load'
        row['Balance Gen MW'] = areas.loc[~is_load, 'Balance MW'].sum()
        row['Balance Load MW'] = areas.loc[is_load, 'Balance MW'].sum()
        err = areas['Error MW'].abs()
        if err.notna().any():
            worst = areas.loc[err.idxmax()]
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

def merge(out_dir: Path) -> Path:
    """
    Merges every <hour>_Log.xlsx in out_dir. Returns the workbook path ('03 Merge Reports.xlsx', beside out_dir).
    Sheets: `hours` (one row per hour), then each of MERGED_SHEETS with an Hour column first. A merged sheet over Excel's
    row limit is written as a CSV beside the workbook instead, and printed.
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

if __name__ == '__main__':
    merge(Path(__file__).parent / 'Output')
