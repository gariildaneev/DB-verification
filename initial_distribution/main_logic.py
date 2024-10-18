import xlsxwriter
import pandas as pd
from collections import deque
from .distribution_utils import write_values, write_module_headers, handle_module_overflow, fa_wise_distribution
from distribution_start import max_signals, fa_order

# Process DI and DO values
def process_discrete_values(db1, conn_diagram, num_DI, num_DO, max_modules, worksheet, workbook, current_row, current_col, current_module, current_section, sections_per_cabinet, cabinet_num):
    db1.FA.astype(int)
    
    db1['FA'] = pd.Categorical(db1['FA'], categories=fa_order, ordered=True)
    db1_sorted = db1.sort_values(by=['CONNECTION', 'FA', 'ID']).reset_index(drop=True)
    di_counter = 1
    do_counter = 1
    di_completed = deque()
    global do_next
    do_next = False
    current_di_connection = None
    current_do_connection = None
    current_di_fa = None
    current_do_fa = None

    temp_kks_di = None
    temp_kks_do = None

    for idx, row in db1_sorted.iterrows():
        kks = row['KKS']
        connection = row['CONNECTION']
        fa = row['FA']

        if pd.isna(kks) or kks == '' or pd.isna(connection) or connection == '':
            continue

        connection_rows = conn_diagram[conn_diagram['CONNECTION'] == connection]
        if connection_rows.empty:
            continue

        di_values = connection_rows['DI'].dropna().unique()
        do_values = connection_rows['DO'].dropna().unique()

        if di_values.size == 0 and do_values.size == 0:
            analog_connection.append((kks, connection, fa))
            continue

        temp_kks_di = (kks, connection, di_values, do_values, fa)

        current_col, current_module, current_row, worksheet, cabinet_num, current_section, current_di_fa, di_counter, do_next = fa_wise_distribution(current_di_fa, fa, di_counter, num_DI, worksheet, workbook, current_row, current_col, current_module, 'DI', cabinet_num, current_section)


        if current_di_connection != connection:
            current_di_connection = connection
            if di_counter > 1:
                if not do_next:
                    temp_kks_di = None
                write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DI')
                while di_counter <= num_DI:
                    write_values(worksheet, workbook, current_row + di_counter, current_col, '', '', '', 'DI')
                    di_counter += 1
                di_counter = 1
                do_next = True
                current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DI', sections_per_cabinet, cabinet_num, current_section)



        if do_next and di_values.size > 0 and di_counter + len(di_values) - 1 > num_DI:
            write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DI')

            while di_counter <= num_DI:
                write_values(worksheet, workbook, current_row + di_counter, current_col, '', '', '', 'DI')
                di_counter += 1

            di_counter = 1
            current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DI', sections_per_cabinet, cabinet_num, current_section)

        elif current_module % 3 != 0 and di_values.size > 0:
            di_completed.append((kks, connection, do_values, fa))
            temp_kks_di = None
            for di_value in di_values:
                if di_counter > num_DI:
                    write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DI')
                    di_counter = 1
                    do_next = True
                    current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DI', sections_per_cabinet, cabinet_num, current_section)


                write_values(worksheet, workbook, current_row + di_counter, current_col, di_value, kks, connection, 'DI', fa)
                di_counter += 1

        if current_module % 3 == 0:
            while di_completed and do_next:

                do_kks, do_connection, do_values, do_fa = di_completed.popleft()
                di_completed.appendleft((do_kks, do_connection, do_values, do_fa))

                current_col, current_module, current_row, worksheet, cabinet_num, current_section, current_do_fa, do_counter, do_next = fa_wise_distribution(current_do_fa, do_fa, do_counter, num_DO, worksheet, workbook, current_row, current_col, current_module, 'DO', cabinet_num, current_section)

                if current_do_connection != do_connection:
                    current_do_connection = do_connection
                    if do_counter > 1:
                        write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DO')
                        while do_counter <= num_DO:
                            write_values(worksheet, workbook, current_row + do_counter, current_col, '', '', '', 'DO')
                            do_counter += 1
                        do_counter = 1
                        do_next = False
                        current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DO', sections_per_cabinet, cabinet_num, current_section)


                if do_next and len(do_values) > 0 and do_counter + len(do_values) - 1 > num_DO:
                    write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DO')

                    while do_counter <= num_DO and do_next:
                        write_values(worksheet, workbook, current_row + do_counter, current_col, '', '', '', 'DO')
                        do_counter += 1
                    do_next = False
                    do_counter = 1
                    current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DO', sections_per_cabinet, cabinet_num, current_section)
                    break

                if do_next and len(do_values) > 0:
                    di_completed.popleft()
                    for do_value in do_values:
                        write_values(worksheet, workbook, current_row + do_counter, current_col, do_value, do_kks, do_connection, 'DO', do_fa)
                        do_counter += 1
                        if do_counter > num_DO:
                            write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DO')
                            do_next = False
                            current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DO', sections_per_cabinet, cabinet_num, current_section)
                            do_counter = 1


        # ДОЗАПОЛНЕНИЕ DO МОДУЛЯ В СЛУЧАЕ, ЕСЛИ DO-ЗНАЧЕНИЯ ЗАКОНЧИЛИСЬ
        if do_counter > 1:
            write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DO')
            while do_counter <= num_DO and do_next:
                write_values(worksheet, workbook, current_row + do_counter, current_col, '', '', '', 'DO')
                do_counter += 1
            do_next = False
            do_counter = 1
            current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DO', sections_per_cabinet, cabinet_num, current_section)

        # ПРОДОЛЖЕНИЕ МОДУЛЯ DI(ПОСЛЕ DO МОДУЛЯ) СО ЗНАЧЕНИЙ, КОТОРЫМ НЕ ХВАТИЛО МЕСТА В ПРЕДЫДУЩЕМ МОДУЛЕ
        if temp_kks_di:
            kks, connection, di_values, do_values, fa = temp_kks_di
            temp_kks_di = None
            di_completed.append((kks, connection, do_values, fa))
            for di_value in di_values:
                write_values(worksheet, workbook, current_row + di_counter, current_col, di_value, kks, connection, 'DI', fa)
                di_counter += 1


    if di_counter > 1:
        write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DI')
        while di_counter <= num_DI:
            write_values(worksheet, workbook, current_row + di_counter, current_col, '', '', '', 'DI')
            di_counter += 1
        current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DI', sections_per_cabinet, cabinet_num, current_section)

    while di_completed:
        do_kks, do_connection, do_values, do_fa = di_completed.popleft()

        current_col, current_module, current_row, worksheet, cabinet_num, current_section, current_do_fa, do_counter, do_next = fa_wise_distribution(current_do_fa, do_fa, do_counter, num_DO, worksheet, workbook, current_row, current_col, current_module, 'DO', cabinet_num, current_section)

        if current_do_connection != do_connection:
            current_do_connection = do_connection
            if do_counter > 1:
                write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DO')
                while do_counter <= num_DO:
                    write_values(worksheet, workbook, current_row + do_counter, current_col, '', '', '', 'DO')
                    do_counter += 1
                do_counter = 1
                current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DO', sections_per_cabinet, cabinet_num, current_section)


        if len(do_values) > 0:
            for do_value in do_values:
                if do_counter > num_DO:
                    write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DO')
                    current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DO', sections_per_cabinet, cabinet_num, current_section)
                    do_counter = 1

                write_values(worksheet, workbook, current_row + do_counter, current_col, do_value, do_kks, do_connection, 'DO', do_fa)
                do_counter += 1

    if do_counter > 1:
        write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'DO')
        while do_counter <= num_DO:
            write_values(worksheet, workbook, current_row + do_counter, current_col, '', '', '', 'DO')
            do_counter += 1
        current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'DO', sections_per_cabinet, cabinet_num, current_section)

    return current_col, current_module, current_section, current_row, worksheet, cabinet_num

# Process AI and AO values
def process_analog_values(analog_connection, conn_diagram, num_AI, num_AO, max_modules, worksheet, workbook, current_row, current_col, current_module, current_section, sections_per_cabinet, cabinet_num):
    ai_counter = 1
    ao_counter = 1
    ao_storage = []
    current_connection = None
    current_ai_fa = None
    current_ao_fa = None

    for kks, connection, fa in analog_connection:
        connection_rows = conn_diagram[conn_diagram['CONNECTION'] == connection]
        ai_values = connection_rows['AI'].dropna().tolist()
        ao_values = connection_rows['AO'].dropna().tolist()

        if len(ao_values) > 0:
            ao_storage.append((kks, connection, ao_values, fa))
        #FIX THAT

        current_col, current_module, current_row, worksheet, cabinet_num, current_section, current_ai_fa, ai_counter, do_next = fa_wise_distribution(current_ai_fa, fa, ai_counter, num_AI, worksheet, workbook, current_row, current_col, current_module, 'AI', cabinet_num, current_section)

        if current_connection != connection:
            current_connection = connection
            if ai_counter > 1:
                write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'AI')
                while ai_counter <= num_AI:
                    write_values(worksheet, workbook, current_row + ai_counter, current_col, '', '', '', 'AI')
                    ai_counter += 1
                ai_counter = 1
                current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'AI', sections_per_cabinet, cabinet_num, current_section)

        if len(ai_values) > 0:
            for ai_value in ai_values:
                if ai_counter > num_AI:
                    write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'AI')
                    ai_counter = 1
                    current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'AI', sections_per_cabinet, cabinet_num, current_section)

                write_values(worksheet, workbook, current_row + ai_counter, current_col, ai_value, kks, connection, 'AI', fa)
                ai_counter += 1

    if ai_counter > 1:
        write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'AI')
        while ai_counter <= num_AI:
            write_values(worksheet, workbook, current_row + ai_counter, current_col, '', '', '', 'AI')
            ai_counter += 1
        current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'AI', sections_per_cabinet, cabinet_num, current_section)

    for kks, connection, ao_values, ao_fa in ao_storage:

        current_col, current_module, current_row, worksheet, cabinet_num, current_section, current_ao_fa, ao_counter, do_next = fa_wise_distribution(current_ao_fa, ao_fa, ao_counter, num_AO, worksheet, workbook, current_row, current_col, current_module, 'AO', cabinet_num, current_section)

        if current_connection != connection:
            current_connection = connection
            if ao_counter > 1:
                write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'AO')
                while ao_counter <= num_AO:
                    write_values(worksheet, workbook, current_row + ao_counter, current_col, '', '', '', 'AO')
                    ao_counter += 1
                ao_counter = 1
                current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'AO', sections_per_cabinet, cabinet_num, current_section)

        for ao_value in ao_values:
            if ao_counter > num_AO:
                write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'AO')
                ao_counter = 1
                current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'AO', sections_per_cabinet, cabinet_num, current_section)

            write_values(worksheet, workbook, current_row + ao_counter, current_col, ao_value, kks, connection, 'AO', ao_fa)
            ao_counter += 1

    if ao_counter > 1:
        write_module_headers(worksheet, workbook, current_row, current_col, current_module, 'AO')
        while ao_counter <= num_AO:
            write_values(worksheet, workbook, current_row + ao_counter, current_col, '', '', '', 'AO')
            ao_counter += 1
        current_col, current_module, current_row, worksheet, cabinet_num, current_section = handle_module_overflow(current_col, current_module, max_modules, current_row, max_signals, worksheet, workbook, 'AO', sections_per_cabinet, cabinet_num, current_section)

    return current_col, current_module, current_section, current_row, worksheet, cabinet_num
