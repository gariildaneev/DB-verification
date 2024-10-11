import openpyxl
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

# Функция предварительной проверки базы данных для сравнения отчетов
def pre_comparison_check(df, file_name):
    kks_column = df['KKS']
    cyrillic_errors = kks_column.apply(contains_cyrillic).any()
    duplicate_errors = kks_column.duplicated().any()

    if cyrillic_errors and duplicate_errors:
        raise ValueError(f"В файле {file_name} в поле KKS обнаружена кириллица и дубликаты, проверьте базу данных.")
    elif cyrillic_errors:
        raise ValueError(f"В файле {file_name} в поле KKS обнаружена кириллица, проверьте базу данных.")
    elif duplicate_errors:
        raise ValueError(f"В файле {file_name} в поле KKS обнаружены дубликаты, проверьте базу данных.")
