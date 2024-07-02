import re
import xlsxwriter

def contains_cyrillic(text):
    return bool(re.search('[а-яА-Я]', text))

def highlight_cyrillic(workbook, worksheet, cyrillic_rows):
    bold_format = workbook.add_format({'bold': True})
    for row_num, row in enumerate(cyrillic_rows.itertuples(), start=1):
        for col_num, cell_value in enumerate(row[1:], start=0):
            if col_num == cyrillic_rows.columns.get_loc('KKS'):
                highlighted_value = ''.join([f'*{char}*' if contains_cyrillic(char) else char for char in cell_value])
                worksheet.write_rich_string(row_num, col_num, *parse_highlighted_text(highlighted_value, bold_format))
            else:
                worksheet.write(row_num, col_num, cell_value)

def parse_highlighted_text(text, format):
    parts = text.split('*')
    formatted_text = []
    for i, part in enumerate(parts):
        if i % 2 == 1:  # Odd index parts are within the * * delimiters
            formatted_text.append(format)
            formatted_text.append(part)
        else:
            formatted_text.append(part)
    return formatted_text
