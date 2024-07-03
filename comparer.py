import pandas as pd

def compare_reports(file1, file2, output_file):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    if 'KKS' not in df1.columns or 'KKS' not in df2.columns:
        raise ValueError("Оба файла должны содержать колонку 'KKS'")

    df1.set_index('KKS', inplace=True)
    df2.set_index('KKS', inplace=True)

    changes = []

    for kks in df1.index:
        if kks in df2.index:
            row1 = df1.loc[kks]
            row2 = df2.loc[kks]
            row_changes = {'KKS': kks}
            for col in df1.columns:
                if col in df2.columns and not pd.isna(row1[col]) and not pd.isna(row2[col]):
                    if row1[col] != row2[col]:
                        row_changes[col] = row1[col]
                        row_changes[f"{col}_new"] = row2[col]
            if len(row_changes) > 1:
                changes.append(row_changes)

    changes_df = pd.DataFrame(changes)

    with xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True}) as workbook:
        ws_changes = workbook.add_worksheet("Отчет о изменениях")

        for c_idx, column in enumerate(changes_df.columns):
            ws_changes.write(0, c_idx, column)

        for r_idx, row in enumerate(changes_df.itertuples(), start=1):
            for c_idx, value in enumerate(row[1:], start=0):
                if pd.isna(value) or value in [float('nan'), float('inf'), float('-inf')]:
                    ws_changes.write(r_idx, c_idx, "")
                elif isinstance(value, (int, float)):
                    ws_changes.write_number(r_idx, c_idx, value)
                else:
                    ws_changes.write(r_idx, c_idx, value)
