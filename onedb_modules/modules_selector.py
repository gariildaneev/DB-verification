import pandas as pd
import xlsxwriter
from onedb_modules.duplicate_module import validate_duplicates
from onedb_modules.cyrillic_module import validate_cyrillic
from onedb_modules.connection_module import validate_connection
from onedb_modules.object_type_module import validate_object_type
from onedb_modules.connection_analytics import analyze_connection

def start_check_process(input_file, output_file, check_duplicates=True, check_cyrillic=True, check_connection=True, check_object_type=True, check_connection_analitycs=True):
    df = pd.read_excel(input_file)
    workbook = xlsxwriter.Workbook(output_file)

    if check_duplicates:
        validate_duplicates(df, workbook)

    if check_cyrillic:
        validate_cyrillic(df, workbook)

    if check_connection:
        validate_connection(df, workbook)

    if check_connection_analitycs:
        analyze_connection(df, workbook)

    if check_object_type:
        validate_object_type(df, workbook)

    workbook.close()
