import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from .distribution_utils import get_unique_fa_values
from .distribution_start import distribution_start
from gui import output, conn_diagram

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
        "Кол-во каналов во входном Дискретном модуле (DI):",
        "Кол-во каналов в выходном Дискретном модуле (DO):",
        "Кол-во каналов в выходном Аналоговом модуле модуле (АI):",
        "Кол-во каналов во входном Аналоговом модуле модуле (АO):",
        "Кол-во модулей в рейке:",
        "Кол-во реек в шкафу:"
    ]
    
    # Clear the window for integer input fields
    for widget in frame.winfo_children():
        widget.destroy()

    # Create IntVar objects to store input values
    global user_ints
    user_ints = [tk.IntVar() for _ in range(6)]  # Create 6 IntVars

    for i in range(6):
        label = ttk.Label(frame, text=custom_texts[i])
        label.grid(row=i, column=0, padx=5, pady=5)

        # Create an entry field for each integer input, bound to an IntVar
        entry = ttk.Entry(frame, textvariable=user_ints[i])
        entry.grid(row=i, column=1, padx=5, pady=5)

    # Button to confirm input values and proceed
    submit_button = ttk.Button(frame, text="Подтвердить ввод", command=lambda: display_info(frame))
    submit_button.grid(row=6, columnspan=2, pady=10)

def display_info(frame):
    global user_ints, db, all_fa_values
    
    # Read the integer values from the input fields before clearing the frame
    try:
        input_values = [var.get() for var in user_ints]  # Get values from IntVar
    except ValueError:
        messagebox.showerror("Invalid Input", "Please enter valid integers.")
        return

    # Clear the window to display the next step
    for widget in frame.winfo_children():
        widget.destroy()

    # Get FA values from the database
    all_fa_values = get_unique_fa_values(db)
    
    # Display the retrieved FA values
    info_label = ttk.Label(frame, text=f"FA, присутствующие в базе данных: {all_fa_values}")
    info_label.grid(row=0, columnspan=2, padx=10, pady=5)

    # Ask for string input for FA grouping
    string_input_label = ttk.Label(frame, text="Введите FA, которые могут находиться в одном модуле, через запятую")
    string_input_label.grid(row=1, column=0, padx=10, pady=5)

    global string_input
    string_input = ttk.Entry(frame)
    string_input.grid(row=1, column=1, padx=10, pady=5)

    # Show the 'Process' button
    process_button = ttk.Button(frame, text="Магия начинается здесь", command=lambda: process_data(input_values))
    process_button.grid(row=2, columnspan=2, pady=10)

def process_data(user_ints):
    global string_input, db, conn_diagram, output, all_fa_values
    
    # Get the string input for FA grouping
    fa_rules = string_input.get()

    # Unpack the integers for the distribution process
    try:
        num_DI, num_DO, num_AI, num_AO, max_modules, sections_per_cabinet = user_ints

        # Run the distribution process
        distribution_start(db, conn_diagram, output, fa_rules, all_fa_values, num_DI, num_DO, num_AI, num_AO, max_modules, sections_per_cabinet)
        
        # Show success message after processing
        messagebox.showinfo("Успех", "Отчет о сравнении успешно создан!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
