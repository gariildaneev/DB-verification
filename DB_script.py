import pandas as pd
import xlsxwriter
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

def contains_cyrillic(text):
    return bool(re.search('[а-яА-Я]', text))

def highlight_cyrillic(text):
    highlighted = []
    for char in text:
        if contains_cyrillic(char):
            highlighted.append((char, True))
        else:
            highlighted.append((char, False))
    return highlighted

def write_highlighted_text(ws, row, col, text, highlight_format):
    highlighted_text = highlight_cyrillic(text)
    for i, (char, is_cyrillic) in enumerate(highlighted_text):
        if is_cyrillic:
            ws.write(row, col + i, char, highlight_format)
        else:
            ws.write(row, col + i, char)

def validate_kks(input_file, output_file, check_unique, check_cyrillic):
    df = pd.read_excel(input_file)

    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    workbook = writer.book

    if check_unique:
        duplicates = df[df.duplicated(subset=['KKS'], keep=False)]
        duplicates.to_excel(writer, sheet_name='Отчет о дубликатах', index=False)
        worksheet_duplicates = writer.sheets['Отчет о дубликатах']
        worksheet_duplicates.write(0, 0, "Значение KKS не уникально")

    if check_cyrillic:
        cyrillic_rows = df[df['KKS'].apply(contains_cyrillic)]
        cyrillic_rows.to_excel(writer, sheet_name='Отчет о кириллице', index=False)
        worksheet_cyrillic = writer.sheets['Отчет о кириллице']
        worksheet_cyrillic.write(0, 0, "Значение KKS содержит кириллицу")

        highlight_format = workbook.add_format({'font_color': 'red', 'bg_color': 'yellow'})

        for row_num, row_data in enumerate(cyrillic_rows.values, start=1):
            for col_num, cell_value in enumerate(row_data):
                if col_num == 2:  # Предполагаем, что KKS в первом столбце
                    write_highlighted_text(worksheet_cyrillic, row_num + 1, col_num, str(cell_value), highlight_format)
                else:
                    worksheet_cyrillic.write(row_num + 1, col_num, cell_value)

    writer.save()

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
                validate_kks(input_file, output_file, var_unique.get(), var_cyrillic.get())
                messagebox.showinfo("Успех", "Отчет успешно создан!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

def create_gui():
    root = tk.Tk()
    root.title("Проверка базы данных")
    root.geometry("300x200")

    # Установка иконки для основного окна
    root.iconbitmap('./_internal/assets/icon.ico')

    global var_unique, var_cyrillic
    var_unique = tk.BooleanVar()
    var_cyrillic = tk.BooleanVar()

    chk_unique = tk.Checkbutton(root, text="Проверка на уникальность", variable=var_unique)
    chk_cyrillic = tk.Checkbutton(root, text="Проверка на кириллицу", variable=var_cyrillic)

    chk_unique.pack()
    chk_cyrillic.pack()

    btn_select_file = tk.Button(root, text="Выбрать файл", command=select_file)
    btn_select_file.pack(expand=True)

    root.mainloop()

if __name__ == "__main__":
    create_gui()


