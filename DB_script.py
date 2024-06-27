import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
import re
import argparse

def contains_cyrillic(text):
    return bool(re.search('[а-яА-Я]', text))

def main(input_file, output_file):
    # Чтение данных из Excel файла
    df = pd.read_excel(input_file)

    # Проверка уникальности значений в столбце "KKS"
    duplicates = df[df.duplicated(subset=['KKS'], keep=False)]

    # Проверка на наличие кириллицы в столбце "KKS"
    cyrillic_rows = df[df['KKS'].apply(contains_cyrillic)]

    # Создание нового Excel файла и листа
    wb = Workbook()

    # Лист для дубликатов
    ws_duplicates = wb.active
    ws_duplicates.title = "Отчет о дубликатах"

    # Добавление заголовка
    ws_duplicates['A1'] = "Значение KKS не уникально"
    ws_duplicates['A1'].font = Font(size=14, bold=True)

    # Запись данных о дубликатах
    for r_idx, r in enumerate(dataframe_to_rows(duplicates, index=False, header=True), start=3):
        for c_idx, value in enumerate(r, start=1):
            ws_duplicates.cell(row=r_idx, column=c_idx, value=value)

    # Лист для строк с кириллицей
    ws_cyrillic = wb.create_sheet(title="Отчет о кириллице")

    # Добавление заголовка
    ws_cyrillic['A1'] = "Значение KKS содержит кириллицу"
    ws_cyrillic['A1'].font = Font(size=14, bold=True)

    # Запись данных о строках с кириллицей
    for r_idx, r in enumerate(dataframe_to_rows(cyrillic_rows, index=False, header=True), start=3):
        for c_idx, value in enumerate(r, start=1):
            ws_cyrillic.cell(row=r_idx, column=c_idx, value=value)

    # Сохранение файла
    wb.save(output_file)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Validate KKS column in Excel file.")
    parser.add_argument("input_file", help="Path to the input Excel file")
    parser.add_argument("output_file", help="Path to the output Excel report file")
    args = parser.parse_args()
    
    main(args.input_file, args.output_file)

