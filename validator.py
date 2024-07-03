import pandas as pd
import xlsxwriter
import re
from utils import contains_cyrillic, highlight_cyrillic

def validate_kks(input_file, output_file, check_duplicates=True, check_cyrillic=True):
    df = pd.read_excel(input_file)
    duplicates = df[df.duplicated(subset=['KKS'], keep=False)]
    cyrillic_rows = df[df['KKS'].apply(contains_cyrillic)]

    workbook = xlsxwriter.Workbook(output_file)

    if check_duplicates:
        ws_duplicates = workbook.add_worksheet("Отчет о дубликатах")
        ws_duplicates.write(0, 0, "Значение KKS не уникально")
        for c_idx, col in enumerate(df.columns):
            ws_duplicates.write(1, c_idx, col)
        for r_idx, row in enumerate(duplicates.itertuples(), start=2):
            for c_idx, value in enumerate(row[1:], start=0):
                if pd.isna(value) or value in [float('nan'), float('inf'), float('-inf')]:
                    ws_duplicates.write(r_idx, c_idx, "")
                elif isinstance(value, (int, float)):
                    ws_duplicates.write_number(r_idx, c_idx, value)
                else:
                    ws_duplicates.write(r_idx, c_idx, str(value))

    if check_cyrillic:
        ws_cyrillic = workbook.add_worksheet("Отчет о кириллице")
        ws_cyrillic.write(0, 0, "Значение KKS содержит кириллицу")
        for c_idx, col in enumerate(df.columns):
            ws_cyrillic.write(1, c_idx, col)
        kks_column_index = df.columns.get_loc('KKS')
        for r_idx, row in enumerate(cyrillic_rows.itertuples(), start=2):
            for c_idx, value in enumerate(row[1:], start=0):
                if c_idx == kks_column_index:
                    highlight_cyrillic(ws_cyrillic, r_idx, c_idx, value, workbook)
                else:
                    if pd.isna(value) or value in [float('nan'), float('inf'), float('-inf')]:
                        ws_cyrillic.write(r_idx, c_idx, "")
                    elif isinstance(value, (int, float)):
                        ws_cyrillic.write_number(r_idx, c_idx, value)
                    else:
                        ws_cyrillic.write(r_idx, c_idx, str(value))

    workbook.close()
