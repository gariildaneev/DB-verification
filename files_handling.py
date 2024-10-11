from tkinter import filedialog

def select_files(num_files, titles_files, title_output):
    files = []
    
    # Loop to ask for the specified number of files with custom titles
    for i in range(num_files):
        file = filedialog.askopenfilename(
            title=titles_files[i],
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            files.append(file)
        else:
            # If the user cancels, return None
            return None, None
    
    # Ask for the output file after selecting input files
    output_file = filedialog.asksaveasfilename(
        title=f"Сохранить {title_output} как",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    
    if output_file:
        # Unpack the files and return them alongside the output file
        return (*files, output_file)
    return None, None
