import xlrd, time
import pandas as pd
from os import listdir

# Module Group - Open required workbooks and check against
def module_group():
    from main_workbook import master_files
    from main_workbook import path

    # Dictionary of columns
    columns = {
        0: "NEW/MOD",
        1: "Reason for the change",
        2: "Module Group Code",
        3: "Module Group Description",
        4: "Imp Country",
        5: "Customer Condition",
        6: "Customer Code",
        7: "Container Group Code",
        8: "Shipping Frequency",
        9: "Module Group Area",
        10: "Module Type Code",
        11: "Filling Efficiency (%)",
        12: "Exp Warehouse"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        4: "TTC Contract Master",
        5: "Module Group Master",
        7: "Customer Master",
        8: "Shipping Calendar Master",
        11: "Container Group Master",
        16: "Module Type Master"
    }

    def check_newmod_field(cell_row, cell_col):
        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != x for x in ('NEW', 'MOD')):
            print ('NEW/MOD check --- Fail')
            update_df(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Please check NEW/MOD field for whitespace')

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 13))

        validate_count = 0
        error_fields = []

        # validate if num   ber
        for col_index in working_columns:
            # validate if blank: Fixed (Probably has its own check, unless Y/N)
            if master_files['xl_sheet_main'].cell_value(9, col_index) == '':
                validate_count += 1
                continue

            try:
                # conditional if number
                maximum_length = int(master_files['xl_sheet_main'].cell_value(9, col_index))
                # check if submitted is integer, remove '.0'
                try:
                    if len(str(int(master_files['xl_sheet_main'].cell_value(cell_row, col_index)))) <= maximum_length:
                        validate_count += 1
                        continue
                    else:
                        error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(9, col_index)))
                        continue
                except ValueError:
                    if len(str(master_files['xl_sheet_main'].cell_value(cell_row, col_index))) <= maximum_length:
                        validate_count += 1
                        continue
                    else:
                        error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(9, col_index)))
                        continue
            except ValueError:
                # conditional if date
                if str(master_files['xl_sheet_main'].cell_value(9, col_index)) == 'dd mmm yyyy':
                    try:
                        time.strptime(str(master_files['xl_sheet_main'].cell_value(cell_row, col_index)),"%d %b %Y")
                        validate_count += 1
                        continue
                    except ValueError:
                        if any(col_index == x for x in (9, 14)): # optional column, can be blank
                            if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                                validate_count += 1
                                continue
                            else:
                                error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(9, col_index)))
                                continue
                        else:
                            error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(9, col_index)))
                            continue
                # conditional if Month
                elif str(master_files['xl_sheet_main'].cell_value(9, col_index)) == 'mmm yyyy':
                    try:
                        time.strptime(str(master_files['xl_sheet_main'].cell_value(cell_row, col_index)),"%b %Y")
                        validate_count += 1
                        continue
                    except ValueError:
                        if col_index == 10: # optional column, can be blank
                            if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                                validate_count += 1
                                continue
                            else:
                                error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(9, col_index)))
                                continue
                        else:
                            error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(9, col_index)))
                            continue
                # conditional if integer,integer
                # Only checks total length, not decimal places
                elif master_files['xl_sheet_main'].cell_value(9, col_index).find(',') != -1:
                    char_limit = master_files['xl_sheet_main'].cell_value(9, col_index).split(',')
                    try:
                        if len(str(master_files['xl_sheet_main'].cell_value(cell_row, col_index))) <= (int(char_limit[0]) + int(char_limit[1]) + 1):
                            validate_count += 1
                            continue
                        else:
                            error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(9, col_index)))
                            continue
                    except ValueError:
                        error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(9, col_index)))
                        continue
                else:
                    error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(9, col_index)))
                    continue

        if validate_count == len(working_columns) and len(error_fields) == 0:
            print ('Maximum Length / Date format check --- Pass')
            update_df(new_mod, 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All fields are within maximum characters')
        else:
            for tuple in error_fields:
                print ('Maximum Length / Date format check --- Fail')
                update_df(new_mod, columns[tuple[0]], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, tuple[0]), tuple[1], 'Check format and character length of submitted')

    def check_compulsory_fields(cell_row, new_mod):
        # Hard code compulsory fields
        compulsory_fields = [2, 3, 4, 5, 7, 8, 9, 10, 11, 12]

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', columns[col_index] + ' is a Compulsory Field')

    # Check for duplicate primary keys
    def module_group_duplicate_key(cell_row, cell_col, new_mod):
        module_group_list = []
        for row in range(10, master_files['xl_sheet_main'].nrows):
            module_group_list.append((master_files['xl_sheet_main'].cell_value(row, 0), master_files['xl_sheet_main'].cell_value(row, 2)))

        matches = 0
        for module_group_code in module_group_list:
            # Check if part no same modifier
            if master_files['xl_sheet_main'].cell_value(cell_row, 0) == module_group_code[0]:
                if PRIMARY_KEY_1 == module_group_code[1]:
                    matches += 1

        if matches == 1:
            print ('Duplicate Key check --- Pass (Primary key is unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, matches, 'Primary key is unique in submitted master')
        elif matches >1:
            print ('Duplicate Key check --- Fail (Primary key is not unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, matches, 'Primary key is not unique in submitted master')

    # Module Group Code must be unique and in this format: ImpCtry(2)Number(3), e.g. MY012
    def module_group_code_check(cell_row, cell_col, new_mod):
        comparison_list_1 = []
        for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_5'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) not in comparison_list_1:
            print ('Module Group Code check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, 'NA', 'Module Group Code is unique')
        else:
            backup_row_contents = []
            for row in range(9, selected['backup_5'].sheet_by_index(0).nrows):
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == selected['backup_5'].sheet_by_index(0).cell_value(row, 2):
                    for col in range(0, 13): # Hard Code column range
                        backup_row_contents.append(selected['backup_5'].sheet_by_index(0).cell_value(row, col))

            submitted_row_contents = []
            for col in range(0, 13): # Hard Code column range
                submitted_row_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            discrepancy_reference = []
            for i, cell_value in enumerate(submitted_row_contents):
                if cell_value != backup_row_contents[i] and all(i != x for x in (0, 1, 2)):
                    discrepancy_reference.append(columns[i])

            if len(discrepancy_reference) == 0:
                discrepancy_reference.append('Submitted has no differences from system')

            print ('Module Group Code check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, ', '.join(discrepancy_reference), 'Module Group Code already exists in system')

        import_country = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)
        if any(import_country == x for x in ('I1', 'I2', 'I3')):
            import_country = 'IN'

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[:2] == import_country and len(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[2:]) == 3:
            try:
                int_transform = int(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[2:])
                print ('Module Group Code check 2 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, import_country + ', ' + str(int_transform), 'Module Group Code in correct format and match Imp Country')
            except ValueError:
                print ('Module Group Code check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Incorrect format: ImpCtry(2)Number(3), e.g. MY012')
        else:
            print ('Module Group Code check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Does not match Imp Country or incorrect format: ImpCtry(2)Number(3), e.g. MY012')

    # If customer condition = S, customer code must be input
    # Customer Code must be found in Customer Master for Imp Country
    def module_group_customer_condition(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'S':
            customer_list = []
            for row in range(9, selected['backup_7'].sheet_by_index(0).nrows):
                customer_list.append(selected['backup_7'].sheet_by_index(0).cell_value(row, 2))

            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1) in customer_list and master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) in master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1).split('-')[0]:
                print ('Customer Condition check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 'Customer Code is correctly input, Customer Code found in Customer Master for correct Imp Country')
            else:
                print ('Customer Condition check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 'If Customer Condition = S, Customer Code must be input, found in Customer Master for correct Imp Country')
        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'M':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1) == '':
                print ('Customer Condition check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 'Customer Condition = M, Customer Code is blank')
            else:
                print ('Customer Condition check --- Warning')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 'Customer Condition = M, Customer Code should be blank')
        else:
            print ('Customer Condition check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 'Customer Condition not correctly input')

    # Container Group Code must be found in container group master for Exp Warehouse Code
    def module_group_container_group(cell_row, cell_col, new_mod):
        container_group_list = []
        for row in range(9, selected['backup_11'].sheet_by_index(0).nrows):
            container_group_list.append(selected['backup_11'].sheet_by_index(0).cell_value(row, 2))

        try:
            for row in range(9, additional['TNM_CONTAINER_GROUP'].nrows):
                container_group_list.append(additional['TNM_CONTAINER_GROUP'].cell_value(row, 2))
        except KeyError:
            pass

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in container_group_list:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[:2] == master_files['xl_sheet_main'].cell_value(cell_row, cell_col+5)[:2]:
                print ('Container Group Check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+5), 'Container Group Code found in System/Submitted for exp warehouse code')
            else:
                print ('Container Group Check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+5), 'Container Group Code does not match Exp Warehouse Code')
        else:
            print ('Container Group Check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Container Group Code not found in System/Submitted')

    # Must overlap with at least 1 ETD of Shipping Calendar Frequency
    def module_group_shipping_frequency(cell_row, cell_col, new_mod):
        module_group_ttc_contract = {}
        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            module_group_ttc_contract[selected['backup_0'].sheet_by_index(0).cell_value(row, 7)] = selected['backup_0'].sheet_by_index(0).cell_value(row, 9)

        try:
            for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                module_group_ttc_contract[additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7)] = additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9)
        except KeyError:
            pass

        ttc_contract_route = {}
        for row in range(9, selected['backup_4'].sheet_by_index(0).nrows):
            ttc_contract_route[selected['backup_4'].sheet_by_index(0).cell_value(row, 2)] = selected['backup_4'].sheet_by_index(0).cell_value(row, 14)

        try:
            for row in range(9, additional['TNM_TTC_CONTRACT'].nrows):
                ttc_contract_route[additional['TNM_TTC_CONTRACT'].cell_value(row, 2)] = additional['TNM_TTC_CONTRACT'].cell_value(row, 14)
        except KeyError:
            pass

        shipping_calendar = {}
        for row in range(9, selected['backup_8'].sheet_by_index(0).nrows):
            shipping_calendar[selected['backup_8'].sheet_by_index(0).cell_value(row, 2)] = selected['backup_8'].sheet_by_index(0).cell_value(row, 12)

        try:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == shipping_calendar[ttc_contract_route[module_group_ttc_contract[master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6)]]]:
                print ('Shipping Frequency check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), shipping_calendar[ttc_contract_route[module_group_ttc_contract[master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6)]]], 'Module Group shipping frequency match with shipping frequency of last ETD of Shipping Route')
            else:
                print ('Shipping Frequency check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), shipping_calendar[ttc_contract_route[module_group_ttc_contract[master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6)]]], 'Module Group shipping frequency does not match with shipping frequency of last ETD of Shipping Route, please check if date overlaps with other ETDs')
        except KeyError:
            print ('Shipping Frequency check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'TTC Contract No. not found in System/Submitted')

    # Must be found in Module Type Master
    def module_group_module_type(cell_row, cell_col, new_mod):
        module_type_list = []
        for row in range(9, selected['backup_16'].sheet_by_index(0).nrows):
            module_type_list.append(selected['backup_16'].sheet_by_index(0).cell_value(row, 3))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in module_type_list:
            print ('Module Type check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Module Type found in Module Type Master')
        else:
            print ('Module Type check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Module Type not found in Module Type Master')

    # Must be found in Warehouse Master
    def module_group_warehouse(cell_row, cell_col, new_mod):
        warehouse_list = []
        for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
            warehouse_list.append(selected['backup_5'].sheet_by_index(0).cell_value(row, 12))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in warehouse_list:
            print ('Warehouse Code check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Warehouse Code registered in Module Group Master before')
        else:
            print ('Warehouse Code check --- Warning')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Warehouse Code not registered before, please check Warehouse Code on User Screen')

    # Use colour to test columns to be MOD
    def get_mod_columns(cell_row):
        cols_to_check = []
        for col in range(0, master_files['xl_sheet_main'].ncols):

            # ------- GET COLOUR INDEX ------- #
            xf_list = master_files['xl_workbook'].xf_list[master_files['xl_sheet_main'].cell_xf_index(cell_row, col)]
            cell_font = master_files['xl_workbook'].font_list[xf_list.font_index]
            # ------- GET COLOUR INDEX ------- #

            if all(col != x for x in (0, 1)):
                if cell_font.colour_index == 10:
                    cols_to_check.append(col-2)

        return cols_to_check

    # Check if RED fields are different from system
    # Check if BLACK fields are same as system
    def module_group_mod_reference(cell_row):

        # Get concat key
        module_group_code = str(master_files['xl_sheet_main'].cell_value(cell_row, 2))

        # Extract all backup concat into list
        comparison_list_1 = []
        for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
            comparison_list_1.append((row, str(selected['backup_5'].sheet_by_index(0).cell_value(row, 2))))

        # Find backup row
        backup_row = 0
        for concat_str in comparison_list_1:
            if module_group_code == concat_str[1]:
                backup_row = concat_str[0]

        # If cannot find, return False
        if backup_row == 0:
            print ('MOD Reference check --- Fail')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Cannot find Backup Row in system for MOD part')
            return False

        if backup_row >= 0:
            # Extract contents of backup row
            backup_row_contents = []
            for col in range(0, selected['backup_5'].sheet_by_index(0).ncols):
                backup_row_contents.append(selected['backup_5'].sheet_by_index(0).cell_value(backup_row, col))

        submitted_contents = []
        validate_count = 0
        for col in range(0, 13): # Hard Code MAX COLUMN

            submitted_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            # ------- GET COLOUR INDEX ------- #
            xf_list = master_files['xl_workbook'].xf_list[master_files['xl_sheet_main'].cell_xf_index(cell_row, col)]
            cell_font = master_files['xl_workbook'].font_list[xf_list.font_index]
            # ------- GET COLOUR INDEX ------- #

            if all(col != x for x in (0, 1)):
                # IF RED
                if cell_font.colour_index == 10:
                    try:
                        if float(master_files['xl_sheet_main'].cell_value(cell_row, col)) != float(backup_row_contents[col]):
                            validate_count += 1
                        else:
                            print ('MOD Reference Check --- Fail (%s RED but not MOD)' % columns[col])
                            update_df('MOD', columns[col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, col), backup_row_contents[col], 'Field is indicated as \'TO MOD (RED)\' but identical to system')
                    except ValueError:
                        if master_files['xl_sheet_main'].cell_value(cell_row, col) != backup_row_contents[col]:
                            validate_count += 1
                        elif all(x == '' for x in (master_files['xl_sheet_main'].cell_value(cell_row, col), backup_row_contents[col])):
                            validate_count += 1
                        else:
                            print ('MOD Reference Check --- Fail (%s RED but not MOD)' % columns[col])
                            update_df('MOD', columns[col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, col), backup_row_contents[col], 'Field is indicated as \'TO MOD (RED)\' but identical to system')
                # IF BLACK
                elif any(cell_font.colour_index == x for x in (8, 32767)):
                    if master_files['xl_sheet_main'].cell_value(cell_row, col) == backup_row_contents[col]:
                        validate_count += 1
                    elif master_files['xl_sheet_main'].cell_value(cell_row, col) == '' and backup_row_contents[col] != '':
                        validate_count += 1
                    else:
                        try:
                            if float(master_files['xl_sheet_main'].cell_value(cell_row, col)) == float(backup_row_contents[col]):
                                validate_count += 1
                            else:
                                print ('MOD Reference Check --- Fail (%s BLACK but MOD)' % columns[col])
                                update_df('MOD', columns[col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, col), backup_row_contents[col], 'Field is indicated as \'UNCHANGED (BLACK)\' but different from system')
                        except ValueError:
                            print ('MOD Reference Check --- Fail (%s BLACK but MOD)' % columns[col])
                            update_df('MOD', columns[col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, col), backup_row_contents[col], 'Field is indicated as \'UNCHANGED (BLACK)\' but different from system')

    # Print required masters for checking
    check_dict = {
        'NEW/MOD': [],
        'Field': [],
        'Row': [],
        'Primary Key': [],
        'Primary Key (Alt)': [],
        'Results': [],
        'Submitted': [],
        'Reference': [],
        'Reason': []
    }

    # Update df dictionary row
    def update_df(new_mod, field, row, primary_key, primary_key_alt, results, submitted, reference, reason):
        check_dict['NEW/MOD'].append(new_mod)
        check_dict['Field'].append(field)
        check_dict['Row'].append(row)
        check_dict['Primary Key'].append(primary_key)
        check_dict['Primary Key (Alt)'].append(primary_key_alt)
        check_dict['Results'].append(results)
        check_dict['Submitted'].append(submitted)
        check_dict['Reference'].append(reference)
        check_dict['Reason'].append(reason)

    print ('The following master files are required: ')
    for i, key in enumerate(required):
        print ('%d: %s' % (i, required.get(key)))
    print ('*' * 60)

    if len(listdir(path + '\\2) Backup')) == 0:
        print ('The folder \'2) Backup\' is empty!')
    else:
        # Automatic loading of backup
        backup_files = []
        for filename in listdir(path + '\\2) Backup\\'):
            backup_files.append(filename)

        selected = {}
        for file in backup_files:
            if file.endswith('.xls') or file.endswith('.csv'):
                if file.find('MRS_CustomerContractDetail') != -1:
                    selected['backup_0'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_TTCContract') != -1:
                    selected['backup_4'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ModuleGroup') != -1:
                    selected['backup_5'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_Customer') != -1 and file.find('MRS_CustomerContract') == -1 and file.find('MRS_CustomerContractDetail') == -1 and file.find('MRS_CustomerPartsMaster') == -1:
                    selected['backup_7'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ShippingCalendar') != -1:
                    selected['backup_8'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ContainerGroup') != -1:
                    selected['backup_11'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ModuleType') != -1:
                    selected['backup_16'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)

        if len(selected) == len(required):
            print('Successfully loaded all backup masters')
            print ('-' * 60)
            print ()
        else:
            # Allow partial retrieval for MOD parts, only if no NEW entries
            change_list = []
            for row in range(10, master_files['xl_sheet_main'].nrows):
                change_list.append(master_files['xl_sheet_main'].cell_value(row, 0))

            if (all(fields != 'NEW' for fields in change_list)):
                print ('WARNING: Not all masters found, program might crash if not all masters available')
                input_to_check = input("Please enter 'Y' to continue, any other key to exit: ")
                if input_to_check != 'Y':
                    return
            else:
                print('Not all masters found, exiting application...')
                return

        # List additional submitted masters to load into memory
        additional = {}
        print ('Loading other submitted masters into memory...')
        print ('Retrieved worksheets:')
        for i, sheet in enumerate(master_files['xl_workbook'].sheets()):
            additional[master_files['xl_workbook'].sheet_by_index(i).name] = master_files['xl_workbook'].sheet_by_index(i)
            print ('%s' % master_files['xl_workbook'].sheet_by_index(i).name)
        print ('-' * 60)
        print ()

        # Checkpoints
        for row in range(10, master_files['xl_sheet_main'].nrows):

            # Set constant for primary keys per master check
            PRIMARY_KEY_1 = master_files['xl_sheet_main'].cell_value(row, 2)
            PRIMARY_KEY_2 = 'NA'

            # If blank, skip row
            if master_files['xl_sheet_main'].cell_value(row, 2) == '':
                continue

            # Print Part no. header
            print ('%s: Part No. %s' % (str(master_files['xl_sheet_main'].cell_value(row, 0)), str(master_files['xl_sheet_main'].cell_value(row, 2))))
            print()

            check_newmod_field(row, 0)

            # Conditional for NEW parts
            if str(master_files['xl_sheet_main'].cell_value(row, 0)).strip(' ') == 'NEW':
                check_maximum_length(row, 'NEW')
                check_compulsory_fields(row, 'NEW')
                module_group_duplicate_key(row, 2, 'NEW')
                module_group_code_check(row, 2, 'NEW')
                module_group_customer_condition(row, 5, 'NEW')
                module_group_container_group(row, 7, 'NEW')
                module_group_shipping_frequency(row, 8, 'NEW')
                module_group_module_type(row, 10, 'NEW')
                module_group_warehouse(row, 12, 'NEW')
            # Conditional for MOD parts
            else:
                customer_condition_check_cycle = 0
                cols_to_check = get_mod_columns(row)
                if len(cols_to_check) != 0:
                    print('User wishes to MOD the following columns:')
                    for col in cols_to_check:
                        print('%s: %s' % (columns[col+2], master_files['xl_sheet_main'].cell_value(row, col+2)))
                    print()

                    if module_group_mod_reference(row):
                        check_maximum_length(row, 'MOD')
                        check_compulsory_fields(row, 'MOD')
                        module_group_duplicate_key(row, 2, 'MOD')

                        for col in cols_to_check:
                            # Mod: Module Group Code
                            if col+2 == 2:
                                print ('%s cannot be modded' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'Cannot be modded')
                            # Mod: Customer Condition, Customer Code
                            if (any(col+2 == x for x in (5, 6))):
                                if customer_condition_check_cycle == 0:
                                    module_group_customer_condition(row, 5, 'MOD')
                                    customer_condition_check_cycle += 1
                            # Mod: Container Group Code
                            if col+2 == 7:
                                module_group_container_group(row, 7, 'MOD')
                            # Mod: Shipping Frequency
                            if col+2 == 8:
                                module_group_shipping_frequency(row, 8, 'MOD')
                            # Mod: Module Type Code
                            if col+2 == 10:
                                module_group_module_type(row, 10, 'MOD')
                            # Mod: Exp Warehouse
                            if col+2 == 12:
                                module_group_warehouse(row, 12, 'MOD')
                            # Mod: optional columns
                            if (any(col+2 == x for x in (3, 4, 9, 11))):
                                print ('There is no programmed check for %s' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'No programmed check')

                else:
                    print ('There is nothing that is being modded.')
                    update_df('MOD', 'NA', row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Nothing is being modded (No field highlighted RED)')

            print ('-' * 10)
        print('*' * 60)

        # Pandas export to excel
        df = pd.DataFrame(check_dict)
        return df
