import pandas as pd

def validate_object_type(df, workbook):
    object_type_analysis_errors = []
    yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
    
    fields_to_check = ['UNITS', 'IN_LEVEL', 'MAX', 'MIN', 'LA', 'HW', 'LW', 'HA', 'HT', 'LT']
    for index, row in df.iterrows():
        object_type = row['OBJECT_TYPE']
        if object_type in ['AI', 'AO']:
            missing_fields = []
            for field in fields_to_check:
                value = row[field]
                if pd.isna(value) or (isinstance(value, str) and value.strip() == ''):
                    missing_fields.append(field)
            if missing_fields:
                row_dict = row.to_dict()
                row_dict['Missing Fields'] = missing_fields
                object_type_analysis_errors.append(row_dict)

    if object_type_analysis_errors:
        ws_object_type_errors = workbook.add_worksheet("Анализ поля OBJECT_TYPE")
        ws_object_type_errors.write(0, 0, "Ошибки: Поля должны быть заполнены для OBJECT_TYPE 'AI' или 'AO'")
        for c_idx, col in enumerate(df.columns):
            ws_object_type_errors.write(1, c_idx, col)
        ws_object_type_errors.write(1, len(df.columns), 'Missing Fields')

        for r_idx, row in enumerate(object_type_analysis_errors, start=2):
            for c_idx, (col, value) in enumerate(row.items()):
                if col == 'Missing Fields':
                    ws_object_type_errors.write(r_idx, len(df.columns), ', '.join(value))
                else:
                    cell_format = yellow_format if col in row['Missing Fields'] else None
                    ws_object_type_errors.write(r_idx, c_idx, str(value) if pd.notna(value) else "", cell_format)
