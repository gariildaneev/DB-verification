import tkinter as tk
from tkinter import filedialog, messagebox
from validator import validate_kks
import os

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
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'icon.ico')
    root.iconbitmap(icon_path)

    root.mainloop()
