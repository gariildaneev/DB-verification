import pandas as pd
import xlsxwriter
import re
from utils import contains_cyrillic, highlight_cyrillic

def validate_kks(input_file, output_file, check_duplicates=True, check_cyrillic=True, check_connection=True):
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
    if check_connection:
        connection_empty_errors = []
        kks_empty_errors = []
    
        for index, row in df.iterrows():
            kks = row['KKS']
            connection = row['Connection']
    
            kks_filled = pd.notna(kks) and kks.strip() != ''
            connection_filled = pd.notna(connection) and connection.strip() != ''
    
            if kks_filled and not connection_filled:
                row_data = row.to_dict()
                row_data['Error'] = 'Connection is empty'
                connection_empty_errors.append(row_data)
            elif not kks_filled and connection_filled:
                row_data = row.to_dict()
                row_data['Error'] = 'KKS is empty'
                kks_empty_errors.append(row_data)
    
        # Создание DataFrame для ошибок
        connection_empty_df = pd.DataFrame(connection_empty_errors)
        kks_empty_df = pd.DataFrame(kks_empty_errors)
    
        # Запись в Excel
        workbook = xlsxwriter.Workbook(output_file)
        worksheet = workbook.add_worksheet("Errors")
    
        # Заголовок для Connection is empty
        worksheet.write(0, 0, "Errors: Connection is empty")
        start_row = 1
        if not connection_empty_df.empty:
            for c_idx, col in enumerate(connection_empty_df.columns, start=0):
                worksheet.write(start_row, c_idx, col)
    
            for r_idx, row in enumerate(connection_empty_df.itertuples(), start=start_row + 1):
                for c_idx, value in enumerate(row[1:], start=0):
                    worksheet.write(r_idx, c_idx, str(value) if pd.notna(value) else "")
    
            start_row += len(connection_empty_df) + 2  # Оставляем отступ
    
        # Заголовок для KKS is empty
        worksheet.write(start_row, 0, "Errors: KKS is empty")
        start_row += 1
        if not kks_empty_df.empty:
            for c_idx, col in enumerate(kks_empty_df.columns, start=0):
                worksheet.write(start_row, c_idx, col)
    
            for r_idx, row in enumerate(kks_empty_df.itertuples(), start=start_row + 1):
                for c_idx, value in enumerate(row[1:], start=0):
                    worksheet.write(r_idx, c_idx, str(value) if pd.notna(value) else "")

    workbook.close()
