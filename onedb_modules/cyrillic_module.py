import pandas as pd
from utils import contains_cyrillic, highlight_cyrillic

def validate_cyrillic(df, workbook):
    cyrillic_rows = df[df['KKS'].apply(lambda x: contains_cyrillic(str(x)) if pd.notna(x) else False)]
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
