import pandas as pd
import xlsxwriter
from utils import contains_cyrillic, highlight_cyrillic

def compare_reports(file1, file2, output_file):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    df1.set_index('KKS', inplace=True)
    df2.set_index('KKS', inplace=True)

    old_kks = set(df1.index)
    new_kks = set(df2.index)

    removed_kks = old_kks - new_kks
    added_kks = new_kks - old_kks
    common_kks = old_kks & new_kks

    changes = []
    
    # Сравнение строк для совпадающих KKS
    for kks in common_kks:
        row1 = df1.loc[kks]
        row2 = df2.loc[kks]

        change = {'KKS': kks}
        has_changes = False

        for col in df1.columns:
            value1, value2 = row1[col], row2[col]

            # Сравнение значений в ячейках
            if pd.isna(value1) or pd.isna(value2) or value1 != value2:
                change[col] = value1
                change[f'{col}_new'] = value2
                has_changes = True
            else:
                change[col] = value1

        if has_changes:
            changes.append(change)

    changes_df = pd.DataFrame(changes)
    
    # Строки с удаленными KKS
    removed_rows = df1.loc[removed_kks].reset_index()
    removed_rows['Status'] = 'Удаленные KKS'

    # Строки с новыми KKS
    added_rows = df2.loc[added_kks].reset_index()
    added_rows['Status'] = 'Новые KKS'

    workbook = xlsxwriter.Workbook(output_file)
    ws_changes = workbook.add_worksheet("Сравнение отчетов")
    ws_removed = workbook.add_worksheet("Удаленные KKS")
    ws_added = workbook.add_worksheet("Новые KKS")

    # Запись изменений
    for c_idx, col in enumerate(changes_df.columns, start=0):
        ws_changes.write(0, c_idx, col)

    for r_idx, row in enumerate(changes_df.itertuples(), start=1):
        for c_idx, value in enumerate(row[1:], start=0):
            if pd.isna(value) or value in [float('nan'), float('inf'), float('-inf')]:
                ws_changes.write(r_idx, c_idx, "")
            elif isinstance(value, (int, float)):
                ws_changes.write_number(r_idx, c_idx, value)
            else:
                ws_changes.write(r_idx, c_idx, str(value))

    # Запись удаленных KKS
    for c_idx, col in enumerate(removed_rows.columns, start=0):
        ws_removed.write(0, c_idx, col)

    for r_idx, row in enumerate(removed_rows.itertuples(), start=1):
        for c_idx, value in enumerate(row[1:], start=0):
            ws_removed.write(r_idx, c_idx, str(value) if pd.notna(value) else "")

    # Запись новых KKS
    for c_idx, col in enumerate(added_rows.columns, start=0):
        ws_added.write(0, c_idx, col)

    for r_idx, row in enumerate(added_rows.itertuples(), start=1):
        for c_idx, value in enumerate(row[1:], start=0):
            ws_added.write(r_idx, c_idx, str(value) if pd.notna(value) else "")

    workbook.close()
