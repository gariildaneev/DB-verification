import pandas as pd
import xlsxwriter
import re
from utils import contains_cyrillic, highlight_cyrillic

def validate_kks(input_file, output_file, check_duplicates=True, check_cyrillic=True, check_connection=True):
    df = pd.read_excel(input_file)
    duplicates = df[df.duplicated(subset=['KKS'], keep=False)]
    cyrillic_rows = df[df['KKS'].apply(lambda x: contains_cyrillic(str(x)) if pd.notna(x) else False)]

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

    if check_connection:
        connection_empty_errors = []
        kks_empty_errors = []
        object_type_empty_errors = []
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

        for index, row in df.iterrows():
            kks = row['KKS']
            connection = row['CONNECTION']
            object_type = row['OBJECT_TYPE']

            kks_filled = pd.notna(kks) and str(kks).strip() != ''
            connection_filled = pd.notna(connection) and str(connection).strip() != ''
            object_type_filled = pd.notna(object_type) and object_type.strip() != ''

            if kks_filled and not connection_filled:
                connection_empty_errors.append(row)
            elif not kks_filled and connection_filled:
                kks_empty_errors.append(row)
            if connection_filled and not object_type_filled:
                object_type_empty_errors.append(row)

        ws_connection_errors = workbook.add_worksheet("Анализ поля CONNECTION")
        start_row = 0

        # Запись ошибок, где Connection пустое
        if connection_empty_errors:
            ws_connection_errors.write(start_row, 0, "Connection is empty", yellow_format)
            start_row += 1
            for c_idx, col in enumerate(df.columns):
                ws_connection_errors.write(start_row, c_idx, col)
            start_row += 1
            for row in connection_empty_errors:
                for c_idx, value in enumerate(row):
                    ws_connection_errors.write(start_row, c_idx, str(value) if pd.notna(value) else "")
                start_row += 1
            start_row += 2  # Оставляем отступ

        # Запись ошибок, где KKS пустое
        if kks_empty_errors:
            ws_connection_errors.write(start_row, 0, "KKS is empty", yellow_format)
            start_row += 1
            for c_idx, col in enumerate(df.columns):
                ws_connection_errors.write(start_row, c_idx, col)
            start_row += 1
            for row in kks_empty_errors:
                for c_idx, value in enumerate(row):
                    ws_connection_errors.write(start_row, c_idx, str(value) if pd.notna(value) else "")
                start_row += 1
        # Запись ошибок, где Object Type пустое
        if object_type_errors:
            ws_connection_errors.write(start_row, 0, "OBJECT_TYPE is empty", yellow_format)
            start_row += 1
            for c_idx, col in enumerate(df.columns):
                ws_connection_errors.write(start_row, c_idx, col)
            start_row += 1
            for row in object_type_errors:
                for c_idx, value in enumerate(row):
                    ws_connection_errors.write(start_row, c_idx, str(value) if pd.notna(value) else "")
                start_row += 1

    workbook.close()
