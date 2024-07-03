import pandas as pd
import xlsxwriter
from utils import contains_cyrillic, highlight_cyrillic

def compare_reports(file1, file2, output_file):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    
    if 'KKS' not in df1.columns or 'KKS' not in df2.columns:
        raise ValueError("Оба файла должны содержать колонку 'KKS'")
    
    df1.set_index('KKS', inplace=True)
    df2.set_index('KKS', inplace=True)

    common_kks = df1.index.intersection(df2.index)

    changes = []

    for kks in common_kks:
        row1 = df1.loc[kks]
        row2 = df2.loc[kks]

        change = {'KKS': kks}
        for col in df1.columns:
            value1, value2 = row1[col], row2[col]

            # Проверка на равенство значений
            if pd.isna(value1) or pd.isna(value2):
                if not pd.isna(value1) or not pd.isna(value2):
                    change[col] = value1
                    change[f'{col}_new'] = value2
            elif isinstance(value1, (pd.Series, pd.DataFrame)):
                if not value1.equals(value2):
                    change[col] = value1
                    change[f'{col}_new'] = value2
            elif hasattr(value1, 'item') and hasattr(value2, 'item'):
                if value1.item() != value2.item():
                    change[col] = value1
                    change[f'{col}_new'] = value2
            elif hasattr(value1, 'all') and hasattr(value2, 'all'):
                if not value1.all() == value2.all():
                    change[col] = value1
                    change[f'{col}_new'] = value2
            else:
                if value1 != value2:
                    change[col] = value1
                    change[f'{col}_new'] = value2
                else:
                    change[col] = value1

        if len(change) > 1:  # If there are any changes other than 'KKS'
            changes.append(change)

    changes_df = pd.DataFrame(changes)

    workbook = xlsxwriter.Workbook(output_file)
    ws = workbook.add_worksheet("Сравнение отчетов")

    for c_idx, col in enumerate(changes_df.columns, start=0):
        ws.write(0, c_idx, col)

    for r_idx, row in enumerate(changes_df.itertuples(), start=1):
        for c_idx, value in enumerate(row[1:], start=0):
            if pd.isna(value) or value in [float('nan'), float('inf'), float('-inf')]:
                ws.write(r_idx, c_idx, "")
            elif isinstance(value, (int, float)):
                ws.write_number(r_idx, c_idx, value)
            else:
                ws.write(r_idx, c_idx, str(value))

    workbook.close()
