import tkinter as tk
from tkinter import filedialog, messagebox
import os
from validator import validate_kks
from comparer import compare_reports

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
            return input_file, output_file
    return None, None

def select_compare_files():
    file1 = filedialog.askopenfilename(
        title="Выберите первый файл (до изменений)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file1:
        file2 = filedialog.askopenfilename(
            title="Выберите второй файл (после изменений)",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file2:
            output_file = filedialog.asksaveasfilename(
                title="Сохранить отчет как",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if output_file:
                return file1, file2, output_file
    return None, None, None

def create_gui():
    root = tk.Tk()
    root.title("Проверка базы данных")
    root.geometry("300x200")

    check_cyrillic = tk.BooleanVar(value=True)
    check_duplicates = tk.BooleanVar(value=True)
    check_compare = tk.BooleanVar(value=True)

    cb_cyrillic = tk.Checkbutton(root, text="Проверка на кириллицу", variable=check_cyrillic)
    cb_duplicates = tk.Checkbutton(root, text="Проверка на дубликаты", variable=check_duplicates)
    cb_compare_check = tk.Checkbutton(root, text="Сравнение отчетов", variable=check_compare)
    

    cb_cyrillic.pack()
    cb_duplicates.pack()
    cb_compare_check.pack()
    
    def on_process():
        if check_duplicates.get() or check_cyrillic.get():
            input_file, output_file = select_file()
            if input_file and output_file:
                try:
                    validate_kks(input_file, output_file, check_duplicates.get(), check_cyrillic.get())
                    messagebox.showinfo("Успех", "Отчет успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
        if check_compare.get():
            file1, file2, output_file = select_compare_files()
            if file1 and file2 and output_file:
                try:
                    compare_reports(file1, file2, output_file)
                    messagebox.showinfo("Успех", "Отчет о сравнении успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

    btn_process = tk.Button(root, text="Запуск", command=on_process)
    btn_process.pack(expand=True)

    # Установка иконки для основного окна
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'icon.ico')
    root.iconbitmap(icon_path)

    root.mainloop()


