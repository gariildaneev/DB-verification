import tkinter as tk
from tkinter import filedialog, messagebox
import os
from validator import validate_kks

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

    check_cyrillic = tk.BooleanVar(value=True)
    check_duplicates = tk.BooleanVar(value=True)

    cb_cyrillic = tk.Checkbutton(root, text="Проверка на кириллицу", variable=check_cyrillic)
    cb_duplicates = tk.Checkbutton(root, text="Проверка на дубликаты", variable=check_duplicates)
    btn_select_file = tk.Button(root, text="Выбрать файл", command=select_file)

    cb_cyrillic.pack()
    cb_duplicates.pack()
    btn_select_file.pack(expand=True)

    # Установка иконки для основного окна
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'icon.ico')
    root.iconbitmap(icon_path)

    root.mainloop()
