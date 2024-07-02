import re

def contains_cyrillic(text):
    return bool(re.search('[а-яА-Я]', text))

def highlight_cyrillic(worksheet, row, col, text, workbook):
    cyrillic_format = workbook.add_format({'bold': True, 'font_color': 'red'})
    parts = re.split('([а-яА-Я]+)', text)
    cell_format = workbook.add_format()
    cell_content = []

    for part in parts:
        if contains_cyrillic(part):
            cell_content.append(cyrillic_format)
            cell_content.append(part)
        else:
            cell_content.append(cell_format)
            cell_content.append(part)

    worksheet.write_rich_string(row, col, *cell_content)
