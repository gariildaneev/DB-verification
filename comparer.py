import pandas as pd
import xlsxwriter
from utils import contains_cyrillic, highlight_cyrillic

def compare_reports(file1, file2, output_file):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    df1.set_index('KKS', inplace=True)
    df2.set_index('KKS', inplace=True)

    common_kks = df1.index.intersection(df2.index)

    changes = []

    for kks in common_kks:
        row1 = df1.loc[kks]
        row2 = df2.loc[kks]

        if not row1.equals(row2):
            change = { 'KKS': kks }
            for col in df1.columns:
                if row1[col] != row2[col]:
                    change[col] = row1[col]
                    change[f'{col}_new'] = row2[col]
                else:
                    change[col] = row1[col]
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
