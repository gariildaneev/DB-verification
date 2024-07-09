import tkinter as tk
from tkinter import ttk
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
    root.title("Проверка баз данных")
    root.geometry("400x300")

    tab_control = ttk.Notebook(root)
    
    tab_single = ttk.Frame(tab_control)
    tab_compare = ttk.Frame(tab_control)
    
    tab_control.add(tab_single, text='Анализ БД')
    tab_control.add(tab_compare, text='Сравнение двух БД')
    
    tab_control.pack(expand=1, fill='both')

    check_cyrillic = tk.BooleanVar(value=False)
    check_duplicates = tk.BooleanVar(value=False)
    check_connection = tk.BooleanVar(value=False)
    check_compare = tk.BooleanVar(value=False)
    check_object_type = tk.BooleanVar(value=False)

    cb_cyrillic = tk.Checkbutton(tab_single, text="Проверка KKS на кириллицу", variable=check_cyrillic)
    cb_duplicates = tk.Checkbutton(tab_single, text="Проверка KKS на дубликаты", variable=check_duplicates)
    cb_connection = tk.Checkbutton(tab_single, text="Анализ поля 'Connection'", variable=check_connection)
    cb_object_type = tk.Checkbutton(tab_single, text="Анализ поля 'Object_type'", variable=check_object_type)

    cb_cyrillic.pack(anchor='w')
    cb_duplicates.pack(anchor='w')
    cb_connection.pack(anchor='w')
    cb_object_type.pack(anchor='w')

    cb_compare_check = tk.Checkbutton(tab_compare, text="Сравнение двух баз данных", variable=check_compare)
    cb_compare_check.pack(anchor='w')

    def on_process_single_file():
        if check_duplicates.get() or check_cyrillic.get() or check_connection.get() or check_object_type.get():
            input_file, output_file = select_file()
            if input_file and output_file:
                try:
                    validate_kks(input_file, output_file, check_duplicates.get(), check_cyrillic.get(), check_connection.get(), check_object_type.get())
                    messagebox.showinfo("Успех", "Отчет успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

    btn_process_single = tk.Button(tab_single, text="Запуск", command=on_process_single_file)
    btn_process_single.pack(expand=True)

    def on_process_compare_files():
        if check_compare.get():
            file1, file2, output_file = select_compare_files()
            if file1 and file2 and output_file:
                try:
                    compare_reports(file1, file2, output_file)
                    messagebox.showinfo("Успех", "Отчет о сравнении успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

    btn_process_compare = tk.Button(tab_compare, text="Запуск", command=on_process_compare_files)
    btn_process_compare.pack(expand=True)


    # Установка иконки для основного окна
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'icon.ico')
    root.iconbitmap(icon_path)

    root.mainloop()


