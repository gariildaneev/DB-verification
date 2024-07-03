import pandas as pd
import xlsxwriter
from utils import contains_cyrillic, highlight_cyrillic

def validate_kks(input_file, output_file, check_duplicates, check_cyrillic):
    df = pd.read_excel(input_file)
    duplicates = df[df.duplicated(subset=['KKS'], keep=False)] if check_duplicates else pd.DataFrame()
    cyrillic_rows = df[df['KKS'].apply(contains_cyrillic)] if check_cyrillic else pd.DataFrame()

    kks_column_index = df.columns.get_loc("KKS")

    with xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True}) as workbook:
        if check_duplicates:
            ws_duplicates = workbook.add_worksheet("Отчет о дубликатах")
            ws_duplicates.write('A1', "Значение KKS не уникально")
            for r_idx, row in enumerate(duplicates.itertuples(), start=1):
                for c_idx, value in enumerate(row[1:], start=1):
                    if pd.isna(value) or value in [float('nan'), float('inf'), float('-inf')]:
                        ws_duplicates.write(r_idx, c_idx, "")
                    elif isinstance(value, (int, float)):
                        ws_duplicates.write_number(r_idx, c_idx, value)
                    else:
                        ws_duplicates.write(r_idx, c_idx, str(value))

        if check_cyrillic:
            ws_cyrillic = workbook.add_worksheet("Отчет о кириллице")
            ws_cyrillic.write('A1', "Значение KKS содержит кириллицу")
            for r_idx, row in enumerate(cyrillic_rows.itertuples(), start=1):
                for c_idx, value in enumerate(row[1:], start=1):
                    if c_idx - 1 == kks_column_index:  # Adjust for zero-based index
                        highlight_cyrillic(ws_cyrillic, r_idx, c_idx, value, workbook)
                    else:
                        if pd.isna(value) or value in [float('nan'), float('inf'), float('-inf')]:
                            ws_cyrillic.write(r_idx, c_idx, "")
                        elif isinstance(value, (int, float)):
                            ws_cyrillic.write_number(r_idx, c_idx, value)
                        else:
                            ws_cyrillic.write(r_idx, c_idx, str(value))
