import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from .distribution_utils import get_unique_fa_values
from .distribution_start import distribution_start

# Global variables to store files and inputs across functions
db = None
conn_diagram = None
output = None
user_ints = []
string_input = None
all_fa_values = None

def create_input_fields(frame, db_data):
    global db
    db = db_data  # Store the database file globally
    
    custom_texts = [
        "Введите первое целое число:",
        "Введите второе целое число:",
        "Введите третье целое число:",
        "Введите четвертое целое число:",
        "Введите пятое целое число:",
        "Введите шестое целое число:"
    ]
    
    # Clear the window for integer input fields
    for widget in frame.winfo_children():
        widget.destroy()
    
    input_vars = []
    
    for i in range(6):  # Assuming 6 inputs
        label = ttk.Label(frame, text=custom_texts[i])
        label.grid(row=i, column=0, padx=5, pady=5)
        
        # Create an entry field for each integer input
        input_var = tk.IntVar()  # You can use StringVar if you want more flexibility
        entry = ttk.Entry(frame, textvariable=input_var)
        entry.grid(row=i, column=1, padx=5, pady=5)
        
        input_vars.append(input_var)
    
    # Store the input variables globally
    global user_ints
    user_ints = input_vars
    
    submit_button = ttk.Button(frame, text="Подтвердить ввод", command=display_info)
    submit_button.grid(row=6, columnspan=2, pady=10)

def display_info():
    # Get and validate the integers
    try:
        global user_ints
        user_ints = [int(var.get()) for var in user_ints]
    except ValueError:
        messagebox.showerror("Invalid Input", "Please enter valid integers.")
        return

    # Clear the window for displaying information
    for widget in frame.winfo_children():
        widget.destroy()

    # Display 'my information' from xlsx files
    global all_fa_values
    all_fa_values = get_unique_fa_values(db)
    info_label = ttk.Label(frame, text=f"FA, присутствующие в базе данных: {all_fa_values}")
    info_label.grid(row=0, columnspan=2, padx=10, pady=5)

    # Proceed to take string input
    string_input_label = ttk.Label(frame, text="Введите FA, которые могут находиться в одном модуле, через запятую")
    string_input_label.grid(row=1, column=0, padx=10, pady=5)

    global string_input
    string_input = ttk.Entry(frame)
    string_input.grid(row=1, column=1, padx=10, pady=5)

    # Show 'Process' button
    process_button = ttk.Button(frame, text="Магия начинается здесь", command=process_data)
    process_button.grid(row=2, columnspan=2, pady=10)

def process_data():
    # Get the string input
    fa_rules = string_input.get()
    num_DI, num_DO, num_AI, num_AO, max_modules, sections_per_cabinet = user_ints

    # Process the data
    try:
        distribution_start(db, conn_diagram, output, fa_rules, all_fa_values, num_DI, num_DO, num_AI, num_AO, max_modules, sections_per_cabinet)
        messagebox.showinfo("Успех", "Отчет о сравнении успешно создан!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
