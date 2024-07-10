import pandas as pd

def validate_duplicates(df, workbook):
    duplicates = df[df.duplicated(subset=['KKS'], keep=False)]
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
