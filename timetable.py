import pandas as pd
import openpyxl
import json
import numpy as np
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

initial_file_path = 'timetable.xlsx'
unmerged_file_path = 'unmerged_timetable.xlsx'

wb = openpyxl.load_workbook(initial_file_path)
ws = wb.active

merged_ranges = list(ws.merged_cells)
print(f"Merged Cell Ranges: {merged_ranges}")

for merged_range in merged_ranges:
    ws.unmerge_cells(str(merged_range))

wb.save(unmerged_file_path)

df_dict = pd.read_excel(unmerged_file_path, sheet_name=None, engine='openpyxl')

cleaned_dfs = []

for sheet_name, df in df_dict.items():
    print(f"Processing sheet: {sheet_name}")
    
    df.columns = df.iloc[0]
    df = df.drop(0)
    df.reset_index(drop=True, inplace=True)

    df.dropna(how='all', inplace=True)
    df.columns = df.columns.str.strip()

    df['INSTRUCTOR-IN-CHARGE / Instructor'] = df['INSTRUCTOR-IN-CHARGE / Instructor'].astype(str).ffill()
    df['ROOM'] = df['ROOM'].astype(str).ffill()
    df['DAYS & HOURS'] = df['DAYS & HOURS'].astype(str).ffill()

    df['SEC'] = df['SEC'].apply(lambda x: x[0] if isinstance(x, list) else x)
    df['INSTRUCTOR-IN-CHARGE / Instructor'] = df['INSTRUCTOR-IN-CHARGE / Instructor'].apply(lambda x: x[0] if isinstance(x, list) else x)

    df.dropna(subset=['COURSE NO.', 'COURSE TITLE'], how='all', inplace=True)

    df['COURSE TITLE'] = df['COURSE TITLE'].str.strip()
    df['INSTRUCTOR-IN-CHARGE / Instructor'] = df['INSTRUCTOR-IN-CHARGE / Instructor'].str.strip()
    df['SEC'] = df['SEC'].str.strip()

    df_grouped = df.groupby(['COURSE NO.', 'COURSE TITLE']).agg({
        'SEC': 'unique',
        'INSTRUCTOR-IN-CHARGE / Instructor': 'unique',
        'ROOM': 'first',
        'DAYS & HOURS': 'first',
    }).reset_index()

    df_grouped = df_grouped[~df_grouped['COURSE NO.'].str.contains('TOTAL|SUMMARY', na=False)]

    df_grouped.reset_index(drop=True, inplace=True)

    df_grouped['SEC'] = df_grouped['SEC'].apply(lambda x: x.tolist() if isinstance(x, np.ndarray) else x)
    df_grouped['INSTRUCTOR-IN-CHARGE / Instructor'] = df_grouped['INSTRUCTOR-IN-CHARGE / Instructor'].apply(lambda x: x.tolist() if isinstance(x, np.ndarray) else x)

    cleaned_dfs.append(df_grouped)

final_df = pd.concat(cleaned_dfs, ignore_index=True)

final_data = final_df.to_dict(orient='records')

with open('timetable.json', 'w', encoding='utf-8') as f:
    json.dump(final_data, f, ensure_ascii=False, indent=4)

print(final_df.head())

