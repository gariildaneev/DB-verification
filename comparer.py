import pandas as pd
import xlsxwriter

def compare_reports(file1, file2, output_file):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    workbook = xlsxwriter.Workbook(output_file)
    ws_compare = workbook.add_worksheet("Отчет о сравнении")

    # Заголовки столбцов
    headers = list(df1.columns) + [f"{col}_new" for col in df1.columns if col in df2.columns]
    for col_num, header in enumerate(headers):
        ws_compare.write(0, col_num, header)

    row_index = 1
    for idx, row in df1.iterrows():
        kks_value = row['KKS']
        match_row = df2[df2['KKS'] == kks_value]

        if not match_row.empty:
            match_row = match_row.iloc[0]
            for col_num, col_name in enumerate(df1.columns):
                ws_compare.write(row_index, col_num, row[col_name])
                if row[col_name] != match_row[col_name]:
                    ws_compare.write(row_index, col_num + len(df1.columns), match_row[col_name])

            row_index += 1

    workbook.close()
