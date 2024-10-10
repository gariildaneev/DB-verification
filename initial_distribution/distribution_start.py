import xlsxwriter
import pandas
from .distribution_utils import write_section_headers, parse_fa_input
from .main_logic import process_discrete_values, process_analog_values

def distribution_start(db1, conn_diagram, output, fa_rules, all_fa_values, num_DI, num_DO, num_AI, num_AO, max_modules, sections_per_cabinet):
  
  # Initialize workbook and worksheet
  workbook = xlsxwriter.Workbook(output)
  worksheet = workbook.add_worksheet('Cab1')

  current_group = {
    'DI': None,
    'DO': None,
    'AI': None,
    'AO': None
  }
  # Initialize starting positions
  current_row = 0
  current_col = 1
  current_module = 1
  current_section = 'A'
  analog_connection = []
  cabinet_num = 1

  fa_groups, fa_order = parse_fa_input(fa_rules, all_fa_values)
  
  # Write the initial section header
  write_section_headers(worksheet, current_section + 'B', current_row, 0, max(num_DI, num_DO, num_AI, num_AO))
  
  # Process DI and DO values
  current_col, current_module, current_section, current_row, worksheet, cabinet_num = process_discrete_values(db1, conn_diagram, num_DI, num_DO, max_modules, worksheet, workbook, current_row, current_col, current_module, current_section, sections_per_cabinet, cabinet_num)
  
  # Process AI and AO values
  current_col, current_module, current_section, current_row, worksheet, cabinet_num = process_analog_values(analog_connection, conn_diagram, num_AI, num_AO, max_modules, worksheet, workbook, current_row, current_col, current_module, current_section, sections_per_cabinet, cabinet_num)
  
  # Close the workbook
  workbook.close()
