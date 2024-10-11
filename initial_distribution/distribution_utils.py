import xlsxwriter
import pandas as pd

def get_unique_fa_values(db1):
    # Assuming db1['FA'] holds FA values as integers or strings
    return sorted(db1['FA'].unique())  # Return sorted unique FA values

def parse_fa_input(fa_rules, all_fa_values):
    # Parse the input and form fa_groups and fa_order, including ungrouped FAs
    groups = fa_rules.split()  # Split by spaces to separate groups
    fa_groups = []  # This will store the tuple groups
    fa_order = []   # This will store the flattened fa values

    for group in groups:
        # Split each group by commas to get the individual fa values and convert to int
        fa_values = tuple(map(int, group.split(',')))
        fa_groups.append(fa_values)  # Add the tuple group
        fa_order.extend(fa_values)   # Add the flattened values to fa_order

    # Add the FA values that were not grouped
    ungrouped_fas = [fa for fa in all_fa_values if fa not in fa_order]
    for fa in ungrouped_fas:
        fa_groups.append((fa,))  # Add each ungrouped FA as its own tuple
        fa_order.append(fa)

    return fa_groups, fa_order

# Write section headers
def write_section_headers(worksheet, section_prefix, row, col, max_values):
    worksheet.write(row, col, section_prefix)
    for i in range(1, max_values + 1):
        worksheet.write(row + i, col, i)

# Write module headers and apply color formatting
def write_module_headers(worksheet, workbook, row, col, module_number, module_type):
    color_map = {
        'DI': '#FFA500',  # Orange
        'DO': '#B19CD9',  # Light Purple
        'AI': '#FFC0CB',  # Pink
        'AO': '#ADD8E6'   # Light Blue
    }
    cell_format = workbook.add_format({'bg_color': color_map[module_type], 'border': 1, 'bold': True})

    worksheet.write(row, col, f"{module_number:02d} ({module_type})", cell_format)
    worksheet.write(row, col + 1, "ext", cell_format)
    worksheet.write(row, col + 2, "CONNECTION", cell_format)

# Write values and apply color formatting
def write_values(worksheet, workbook, row, col, value, kks, connection, module_type, fa=None):
    color_map = {
        'DI': '#FFA500',  # Orange
        'DO': '#B19CD9',  # Light Purple
        'AI': '#FFC0CB',  # Pink
        'AO': '#ADD8E6'   # Light Blue
    }
    cell_format = workbook.add_format({'bg_color': color_map[module_type], 'border': 1})

    worksheet.write(row, col, kks, cell_format)
    if fa is not None:
        worksheet.write_comment(row, col, 'FA: ' + str(fa), {'font_size': 14})
    worksheet.write(row, col + 1, value, cell_format)
    worksheet.write(row, col + 2, connection, cell_format)

# Handle module overflow and advance to the next section if needed
def handle_module_overflow(current_col, current_module, max_modules, row, max_values, worksheet, workbook, module_type, sections_per_cabinet, cabinet_num, current_section):
    current_col += 3
    current_module += 1
    if current_module > max_modules:
        global do_next
        do_next = False
        current_section = chr(ord(current_section) + 1)
        row += max_values + 1
        current_col = 1
        current_module = 1
        if ord(current_section) - ord('A') >= sections_per_cabinet:
            cabinet_num += 1
            current_section = 'A'
            row = 0
            worksheet = workbook.add_worksheet(f'Cab{cabinet_num}')
        write_section_headers(worksheet, current_section + 'B', row, 0, max_values)
    return current_col, current_module, row, worksheet, cabinet_num, current_section

def fa_wise_distribution(current_fa, fa, counter, num, worksheet, workbook, current_row, current_col, current_module, module_type, cabinet_num, current_section):
    global do_next

    group_found = False
    if current_fa != fa:
        current_fa = fa
        for group in fa_groups:
            if current_fa in group:
                group_found = True
                if current_group[module_type] is None or current_group[module_type] != group:
                    # New group encountered, switch module
                    current_group[module_type] = group
                    if counter > 1:
                        write_module_headers(worksheet, workbook, current_row, current_col, current_module, module_type)
                        while counter <= num:
                            write_values(worksheet, workbook, current_row + counter, current_col, '', '', '', module_type)
                            counter += 1
                        if module_type == 'DI':
                            do_next = True
                        else:
                            do_next = False
                        counter = 1
                        current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max(num_DI, num_DO, num_AI, num_AO), worksheet, workbook, module_type, sections_per_cabinet, cabinet_num, current_section)
                break

        if not group_found:
            raise ValueError(f"FA value {current_fa} not found in user input")
    return current_col, current_module, current_row, worksheet, cabinet_num, current_section, current_fa, counter, do_next
