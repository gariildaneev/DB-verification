import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import xlsxwriter

def contains_cyrillic(text):
    return bool(re.search('[а-яА-Я]', text))

def highlight_cyrillic(workbook, worksheet, cyrillic_rows):
    format_highlight = workbook.add_format({'bg_color': 'yellow'})
    for row_num, row in enumerate(cyrillic_rows.itertuples(), start=1):
        for col_num, cell_value in enumerate(row[1:], start=0):
            if col_num == cyrillic_rows.columns.get_loc('KKS'):
                if contains_cyrillic(cell_value):
                    worksheet.write(row_num, col_num, cell_value, format_highlight)
                else:
                    worksheet.write(row_num, col_num, cell_value)
            else:
                worksheet.write(row_num, col_num, cell_value)

def validate_kks(input_file, output_file, check_cyrillic=True, check_duplicates=True):
    df = pd.read_excel(input_file)
    wb = xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True})
    
    if check_duplicates:
        duplicates = df[df.duplicated(subset=['KKS'], keep=False)]
        ws_duplicates = wb.add_worksheet("Отчет о дубликатах")
        ws_duplicates.write(0, 0, "Значение KKS не уникально")
        for r_idx, row in enumerate(duplicates.itertuples(), start=1):
            for c_idx, value in enumerate(row[1:], start=0):
                ws_duplicates.write(r_idx, c_idx, value)

    if check_cyrillic:
        cyrillic_rows = df[df['KKS'].apply(contains_cyrillic)]
        ws_cyrillic = wb.add_worksheet("Отчет о кириллице")
        ws_cyrillic.write(0, 0, "Значение KKS содержит кириллицу")
        highlight_cyrillic(wb, ws_cyrillic, cyrillic_rows)

    wb.close()

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
    root.geometry("300x200")

    chk_var_cyrillic = tk.BooleanVar()
    chk_var_duplicates = tk.BooleanVar()

    chk_cyrillic = tk.Checkbutton(root, text="Проверка на кириллицу", variable=chk_var_cyrillic)
    chk_duplicates = tk.Checkbutton(root, text="Проверка на дубликаты", variable=chk_var_duplicates)

    chk_cyrillic.pack(anchor='w')
    chk_duplicates.pack(anchor='w')

    def on_select_file():
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
                    validate_kks(input_file, output_file, chk_var_cyrillic.get(), chk_var_duplicates.get())
                    messagebox.showinfo("Успех", "Отчет успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

    btn_select_file = tk.Button(root, text="Выбрать файл", command=on_select_file)
    btn_select_file.pack(expand=True)

    # Установка иконки для основного окна
    root.iconbitmap('./_internal/assets/icon.ico')

    root.mainloop()

if __name__ == "__main__":
    create_gui()

