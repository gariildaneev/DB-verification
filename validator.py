import pandas as pd
import re
import xlsxwriter
from utils import contains_cyrillic, highlight_cyrillic

def validate_kks(input_file, output_file, check_cyrillic=True, check_duplicates=True):
    df = pd.read_excel(input_file)
    wb = xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True})
    
    if check_duplicates:
        duplicates = df[df.duplicated(subset=['KKS'], keep=False)]
        ws_duplicates = wb.add_worksheet("Отчет о дубликатах")
        ws_duplicates.write(0, 0, "Значение KKS не уникально")
        for r_idx, row in enumerate(duplicates.itertuples(), start=1):
            for c_idx, value in enumerate(row[1:], start=0):
                ws_duplicates.write(r_idx, c_idx, value)

    if check_cyrillic:
        cyrillic_rows = df[df['KKS'].apply(contains_cyrillic)]
        ws_cyrillic = wb.add_worksheet("Отчет о кириллице")
        ws_cyrillic.write(0, 0, "Значение KKS содержит кириллицу")
        highlight_cyrillic(wb, ws_cyrillic, cyrillic_rows)

    wb.close()
