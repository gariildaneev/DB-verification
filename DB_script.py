import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
import re
import tkinter as tk
from tkinter import filedialog, messagebox

def contains_cyrillic(text):
    return bool(re.search('[а-яА-Я]', text))

def validate_kks(input_file, output_file):
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

def select_file():
    input_file = filedialog.askopenfilename(
        title="Выберите файл",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if input_file:
        output_file = filedialog.asksaveasfilename(
            title="Сохранить отчет как",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if output_file:
            try:
                validate_kks(input_file, output_file)
                messagebox.showinfo("Успех", "Отчет успешно создан!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

def create_gui():
    root = tk.Tk()
    root.title("Проверка базы данных")
    root.geometry("300x150")

    btn_select_file = tk.Button(root, text="Выбрать файл", command=select_file)
    btn_select_file.pack(expand=True)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
