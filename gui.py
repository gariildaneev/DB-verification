import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sv_ttk
from onedb_modules.modules_selector import start_check_process
from comparer import compare_reports, compare_with_connection_schema
from files_handling import select_files
from initial_distribution.inputs_handling_preprocessing import create_input_fields

def create_gui():
    root = tk.Tk()
    root.title("Проверка баз данных")
    root.geometry("500x400")

    tab_control = ttk.Notebook(root)
    
    tab_single = ttk.Frame(tab_control)
    tab_compare = ttk.Frame(tab_control)
    tab_distribution = ttk.Frame(tab_control)
    
    tab_control.add(tab_single, text='Анализ БД')
    tab_control.add(tab_compare, text='Сравнение двух БД')
    tab_control.add(tab_distribution, text='Первичное распределение сигналов')
    
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
        frame = ttk.Frame(tab)
        checkbox = ttk.Checkbutton(frame, text=text, variable=variable)
        info_button = ttk.Button(frame, text="?", command=show_info(info_text), width=2)
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
            input_file, output_file = select_files(num_files=1, titles_files=["Выберите БД"], titles_output="отчёт")
            if input_file and output_file:
                try:
                    start_check_process(input_file, output_file, check_duplicates.get(), check_cyrillic.get(), check_connection.get(), check_object_type.get(), check_connection_analitycs.get())
                    messagebox.showinfo("Успех", "Отчет успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

    btn_process_single = ttk.Button(tab_single, text="Запуск", command=on_process_single_file)
    btn_process_single.pack(expand=True)

    def on_process_compare_files():
        if check_compare.get():
            file1, file2, output_file = select_files(num_files=2, titles_files=["Выберите БД №1", "Выберите БД №2"], titles_output="отчёт")
            if file1 and file2 and output_file:
                try:
                    compare_reports(file1, file2, output_file)
                    messagebox.showinfo("Успех", "Отчет о сравнении успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
        if check_unknown_connection.get():
            file1, file2, output_file = select_files(num_files=2, titles_files=["Выберите БД", "Выберите Connection diagram"], titles_output="отчёт")
            if file1 and file2 and output_file:
                try:
                    compare_with_connection_schema(file1, file2, output_file)
                    messagebox.showinfo("Успех", "Отчет о сравнении успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

    btn_process_compare = ttk.Button(tab_compare, text="Запуск", command=on_process_compare_files)
    btn_process_compare.pack(expand=True)

    def on_process_initial_distribution():
        file1, file2, output_file = select_files(num_files=2, titles_files=["Выберите БД", "Выберите Connection diagram"], titles_output="первичное распределение сигналов")
        if file1 and file2 and output_file:
                try:
                    db = pd.read_excel(file1)
                    conn_diagram = pd.read_excel(file2)
                    
                    create_input_fields(tab_distribution, db) 
                    
                    messagebox.showinfo("Успех", "Отчет о сравнении успешно создан!")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

    btn_process_initial_distribution = ttk.Button(tab_distribution, text="Запуск", command=on_process_initial_distribution)
    btn_process_initial_distribution.pack(expand=True)
    
    # Установка иконки для основного окна
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'icon.ico')
    root.iconbitmap(icon_path)

    sv_ttk.use_light_theme()
    
    root.mainloop()
