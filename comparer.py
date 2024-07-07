import pandas as pd
import xlsxwriter
from difflib import ndiff
from utils import pre_comparison_check


def compare_reports(file1, file2, output_file):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Проверка структуры таблиц
    if not df1.columns.equals(df2.columns):
        raise ValueError("Проверьте структуру таблицы: заголовки должны совпадать")
    
    # Проверка файлов на дубликаты и кириллицу в KKS
    pre_comparison_check(df1, file1)
    pre_comparison_check(df2, file2)
    
    # Если ошибок нет, выполняем сравнение отчетов
    df1.set_index('KKS', inplace=True)
    df2.set_index('KKS', inplace=True)
    
    old_kks = set(df1.index)
    new_kks = set(df2.index)
    
    removed_kks = old_kks - new_kks
    added_kks = new_kks - old_kks
    common_kks = old_kks & new_kks
    
    changes = []
    
    for kks in common_kks:
        row1 = df1.loc[kks]
        row2 = df2.loc[kks]
    
        change = {'KKS': kks}
        has_changes = False
    
        for col in df1.columns:
    
            value1 = str(row1[col]) if pd.notna(row1[col]) else ""
    
            value2 = str(row2[col]) if pd.notna(row2[col]) else ""
    
            if value1 != value2:
                change[col] = value1
                change[f'{col}_new'] = value2
                has_changes = True
            else:
                change[col] = value1
    
        if has_changes:
            changes.append(change)
    
    
    changes_df = pd.DataFrame(changes)
    
    removed_rows = df1.loc[list(removed_kks)].reset_index()
    removed_rows['Status'] = 'Удаленные KKS'
    
    added_rows = df2.loc[list(added_kks)].reset_index()
    added_rows['Status'] = 'Новые KKS'
    
    workbook = xlsxwriter.Workbook(output_file)
    ws_changes = workbook.add_worksheet("Сравнение БД")
    ws_removed = workbook.add_worksheet("Удаленные KKS")
    ws_added = workbook.add_worksheet("Новые KKS")
    
        # Формат для выделения колонок _new
    yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
    
    # Записываем заголовки в лист сравнения отчетов
    change_columns = []
    for col in changes_df.columns:
        if not col.endswith("_new"):
            change_columns.append(col)
            if f'{col}_new' in changes_df.columns:
                change_columns.append(f'{col}_new')
    
    for c_idx, col in enumerate(change_columns, start=0):
        header_format = yellow_format if col.endswith('_new') else None
        ws_changes.write(0, c_idx, col, header_format)
    
    # Записываем данные в лист сравнения отчетов
    for r_idx, row in enumerate(changes_df.itertuples(), start=1):
        for c_idx, col in enumerate(change_columns, start=0):
            value = getattr(row, col, "")
            cell_format = yellow_format if col.endswith('_new') and value else None
            if pd.isna(value) or value in [float('nan'), float('inf'), float('-inf')]:
                ws_changes.write(r_idx, c_idx, "", cell_format)
            elif isinstance(value, (int, float)):
                ws_changes.write_number(r_idx, c_idx, value, cell_format)
            else:
                ws_changes.write(r_idx, c_idx, str(value), cell_format)
    
    
    for c_idx, col in enumerate(removed_rows.columns, start=0):
        ws_removed.write(0, c_idx, col)
    
    for r_idx, row in enumerate(removed_rows.itertuples(), start=1):
        for c_idx, value in enumerate(row[1:], start=0):
            ws_removed.write(r_idx, c_idx, str(value) if pd.notna(value) else "")
    
    for c_idx, col in enumerate(added_rows.columns, start=0):
        ws_added.write(0, c_idx, col)
    
    for r_idx, row in enumerate(added_rows.itertuples(), start=1):
        for c_idx, value in enumerate(row[1:], start=0):
            ws_added.write(r_idx, c_idx, str(value) if pd.notna(value) else "")
    
    workbook.close()
