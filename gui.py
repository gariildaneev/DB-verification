import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import os
from onedb_modules.modules_selector import start_check_process
from comparer import compare_reports, compare_with_connection_schema

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
        title="Выберите базу данных",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file1:
        file2 = filedialog.askopenfilename(
            title="Выберите БД для сравнения",
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
    check_connection_analitycs = tk.BooleanVar(value=False)
    check_unknown_connection = tk.BooleanVar(value=False)

    def show_info(info_text):
        def _show_info():
            messagebox.showinfo("Информация", info_text)
        return _show_info

    def create_checkbox_with_info(tab, text, variable, info_text):
        frame = tk.Frame(tab)
        checkbox = tk.Checkbutton(frame, text=text, variable=variable)
        info_button = tk.Button(frame, text="?", command=show_info(info_text), width=2)
        checkbox.pack(side="left", anchor='w')
        info_button.pack(side="left", anchor='w')
        frame.pack(anchor='w', pady=2)

    create_checkbox_with_info(tab_single, "Проверка KKS на кириллицу", check_cyrillic, "Проверяет наличие кириллических символов в поле KKS.")
    create_checkbox_with_info(tab_single, "Проверка KKS на дубликаты", check_duplicates, "Проверяет, что все значения KKS уникальны.")
    create_checkbox_with_info(tab_single, "Анализ поля 'Connection'", check_connection, "Анализирует поле 'Connection' и проверяет его на корректность.")
    create_checkbox_with_info(tab_single, "Анализ поля 'Object_type'", check_object_type, "Проверяет поля для OBJECT_TYPE 'AI' или 'AO'.")
    create_checkbox_with_info(tab_single, "CONNECTION-аналитика", check_connection_analitycs, "Показывает статистику по полю 'Connection'.")

    create_checkbox_with_info(tab_compare, "Сравнение двух баз данных", check_compare, "Сравнивает две базы данных и выводит отчет о различиях.")
    create_checkbox_with_info(tab_compare, "Проверка схем подключений", check_unknown_connection, "Сравнение значений поля 'CONNECTION' с базой схем подключения и выявление неизвестных типов подключений.")

    def on_process_single_file():
        if check_duplicates.get() or check_cyrillic.get() or check_connection.get() or check_object_type.get() or check_connection_analitycs.get():
            input_file, output_file = select_file()
            if input_file and output_file:
                try:
                    start_check_process(input_file, output_file, check_duplicates.get(), check_cyrillic.get(), check_connection.get(), check_object_type.get(), check_connection_analitycs.get())
                    messagebox.showinfo("Успех", "Отчет успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

    btn_process_single = tk.Button(tab_single, text="Запуск", command=on_process_single_file)
    btn_process_single.pack(expand=True)

    def on_process_two_files():
        if check_compare.get():
            file1, file2, output_file = select_compare_files()
            if file1 and file2 and output_file:
                try:
                    compare_reports(file1, file2, output_file)
                    messagebox.showinfo("Успех", "Отчет о сравнении успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
        if check_unknown_connection.get():
            file1, file2, output_file = select_compare_files()
            if file1 and file2 and output_file:
                try:
                    compare_with_connection_schema(file1, file2, output_file)
                    messagebox.showinfo("Успех", "Отчет о сравнении успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

    btn_process_compare = tk.Button(tab_compare, text="Запуск", command=on_process_two_files)
    btn_process_compare.pack(expand=True)

    # Установка иконки для основного окна
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'icon.ico')
    root.iconbitmap(icon_path)

    root.mainloop()
