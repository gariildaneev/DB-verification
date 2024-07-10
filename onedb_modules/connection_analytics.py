import pandas as pd

def analyze_connection(df, workbook):
    connection_counts = df[df['KKS'].notna() & df['KKS'].str.strip().astype(bool)]['CONNECTION'].value_counts().reset_index()
    connection_counts.columns = ['Connection', 'Кол-во']
    
    ws_connection_analysis = workbook.add_worksheet("CONNECTION-аналитика")
    border_format = workbook.add_format({'border': 1})
    
    headers = ['Connection', 'Кол-во']
    for c_idx, header in enumerate(headers):
        ws_connection_analysis.write(0, c_idx, header, border_format)
    
    for r_idx, row in connection_counts.iterrows():
        for c_idx, value in enumerate(row):
            ws_connection_analysis.write(r_idx + 1, c_idx, str(value) if pd.notna(value) else "", border_format)
