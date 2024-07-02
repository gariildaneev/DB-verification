import pandas as pd
import xlsxwriter
from utils import contains_cyrillic, highlight_cyrillic

def validate_kks(input_file, output_file):
    df = pd.read_excel(input_file)
    duplicates = df[df.duplicated(subset=['KKS'], keep=False)]
    cyrillic_rows = df[df['KKS'].apply(contains_cyrillic)]

    workbook = xlsxwriter.Workbook(output_file)

    # Sheet for duplicates
    ws_duplicates = workbook.add_worksheet("Отчет о дубликатах")
    ws_duplicates.write('A1', "Значение KKS не уникально", workbook.add_format({'bold': True, 'font_size': 14}))

    for r_idx, row in enumerate(duplicates.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=0):
            if pd.isna(value):
                ws_duplicates.write_blank(r_idx, c_idx, None)
            elif isinstance(value, (int, float)):
                ws_duplicates.write_number(r_idx, c_idx, value)
            else:
                ws_duplicates.write_string(r_idx, c_idx, str(value))

    # Sheet for cyrillic
    ws_cyrillic = workbook.add_worksheet("Отчет о кириллице")
    ws_cyrillic.write('A1', "Значение KKS содержит кириллицу", workbook.add_format({'bold': True, 'font_size': 14}))

    for r_idx, row in enumerate(cyrillic_rows.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=0):
            if pd.isna(value):
                ws_cyrillic.write_blank(r_idx, c_idx, None)
            elif c_idx == cyrillic_rows.columns.get_loc('KKS'):
                highlight_cyrillic(ws_cyrillic, r_idx, c_idx, value, workbook)
            elif isinstance(value, (int, float)):
                ws_cyrillic.write_number(r_idx, c_idx, value)
            else:
                ws_cyrillic.write_string(r_idx, c_idx, str(value))

    workbook.close()
