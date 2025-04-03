import pandas as pd
from pathlib import Path
import Scripts.wpp_lib as wpp_lib

cur_dir = Path(__file__).parent

# Directory and file pattern
input_dir = Path("Output")
excel_files = list(input_dir.glob("*_ScaleLog.xlsx"))
target_fp = cur_dir / "03 Merge Reports.xlsx"
# Dictionary to store DataFrames by sheet name
sheets_data = {}

# Loop through each Excel file
for file in excel_files:
    try:
        # Read all sheets from the workbook
        sheets_dict = pd.read_excel(file, sheet_name=None)  # Returns {sheet_name: DataFrame}

        for sheet_name, df in sheets_dict.items():
            # Insert the filename as the first column
            df.insert(0, 'Filename', file.name)

            # Append to the corresponding sheet in sheets_data
            if sheet_name in sheets_data:
                sheets_data[sheet_name].append(df)
            else:
                sheets_data[sheet_name] = [df]  # Initialize a new list for the sheet

    except Exception as e:
        print(f"Error processing {file.name}: {e}")

# Concatenate all dataframes. 
for sheet_name, df_list in sheets_data.items():
    sheets_data[sheet_name] = pd.concat(df_list, ignore_index=True)

wpp_lib.df_dict_to_excel_workbook(target_fp, sheets_data)

print(f"Concatenation complete. Data saved to '{str(target_fp)}'.")
