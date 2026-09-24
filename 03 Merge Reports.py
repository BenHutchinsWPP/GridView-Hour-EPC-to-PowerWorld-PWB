from collections import defaultdict
from pathlib import Path
import pandas as pd
from Scripts import pw

cur_dir = Path(__file__).parent

# Directory and file pattern
input_dir = cur_dir / "Output"
excel_files = sorted(input_dir.glob("*_ScaleLog.xlsx"))
target_fp = cur_dir / "03 Merge Reports.xlsx"

if len(excel_files) == 0:
    raise SystemExit(f'No *_ScaleLog.xlsx files found in {input_dir}')

# DataFrames by sheet name
sheets_data = defaultdict(list)

for file in excel_files:
    # Read all sheets from the workbook, and insert the filename as the first column.
    for sheet_name, df in pd.read_excel(file, sheet_name=None).items():
        df.insert(0, 'Filename', file.name)
        sheets_data[sheet_name].append(df)

pw.write_excel(target_fp, {name: pd.concat(dfs, ignore_index=True) for name, dfs in sheets_data.items()})

print(f"Concatenation complete. Data saved to '{str(target_fp)}'.")
