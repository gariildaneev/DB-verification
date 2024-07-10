import pandas as pd

def validate_connection(df, workbook):
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
        object_type_filled = pd.notna(object_type) and str(object_type).strip() != ''

        if kks_filled and not connection_filled:
            connection_empty_errors.append(row)
        elif not kks_filled and connection_filled:
            kks_empty_errors.append(row)
        if connection_filled and not object_type_filled:
            object_type_empty_errors.append(row)

    ws_connection_errors = workbook.add_worksheet("Аналитика поля Connection")
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
    if object_type_empty_errors:
        ws_connection_errors.write(start_row, 0, "OBJECT_TYPE is empty", yellow_format)
        start_row += 1
        for c_idx, col in enumerate(df.columns):
            ws_connection_errors.write(start_row, c_idx, col)
        start_row += 1
        for row in object_type_empty_errors:
            for c_idx, value in enumerate(row):
                ws_connection_errors.write(start_row, c_idx, str(value) if pd.notna(value) else "")
            start_row += 1
