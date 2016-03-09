import time
from os import listdir

import xlrd
import pandas as pd

# Build-out - Open required workbooks and check against
def build_out(master_files, path):

    # Dictionary of columns
    columns = {
        0: "NEW/MOD",
        1: "Reason for the change",
        2: "TTC Parts No.",
        3: "Colour Code",
        4: "Customer Contract No.",
        5: "Customer Code",
        6: "Build-out Indicator",
        7: "Build-out Priority",
        8: "Change Type",
        9: "Build-out Date",
        10: "Last Customer Order Month",
        11: "Build-out Stock Quantity",
        12: "Order Lot Flag",
        13: "Flag 2 Parts",
        14: "Last Supplier Delivery Date"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        16: "Build-out Master"
    }

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

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 15))

        validate_count = 0
        error_fields = []

        # validate if num   ber
        for col_index in working_columns:
            # validate if blank: Fixed (Probably has its own check, unless Y/N)
            if master_files['xl_sheet_main'].cell_value(18, col_index) == '':
                validate_count += 1
                continue

            try:
                # conditional if number
                maximum_length = int(master_files['xl_sheet_main'].cell_value(18, col_index))
                # check if submitted is integer, remove '.0'
                try:
                    if len(str(int(master_files['xl_sheet_main'].cell_value(cell_row, col_index)))) <= maximum_length:
                        validate_count += 1
                        continue
                    else:
                        error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(18, col_index)))
                        continue
                except ValueError:
                    if len(str(master_files['xl_sheet_main'].cell_value(cell_row, col_index))) <= maximum_length:
                        validate_count += 1
                        continue
                    else:
                        error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(18, col_index)))
                        continue
            except ValueError:
                # conditional if date
                if str(master_files['xl_sheet_main'].cell_value(18, col_index)) == 'dd mmm yyyy':
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
                                error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(18, col_index)))
                                continue
                        else:
                            error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(18, col_index)))
                            continue
                # conditional if Month
                elif str(master_files['xl_sheet_main'].cell_value(18, col_index)) == 'mmm yyyy':
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
                                error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(18, col_index)))
                                continue
                        else:
                            error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(18, col_index)))
                            continue
                # conditional if integer,integer
                # Only checks total length, not decimal places
                elif master_files['xl_sheet_main'].cell_value(18, col_index).find(',') != -1:
                    char_limit = master_files['xl_sheet_main'].cell_value(18, col_index).split(',')
                    try:
                        if len(str(master_files['xl_sheet_main'].cell_value(cell_row, col_index))) <= (int(char_limit[0]) + int(char_limit[1]) + 1):
                            validate_count += 1
                            continue
                        else:
                            error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(18, col_index)))
                            continue
                    except ValueError:
                        error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(18, col_index)))
                        continue
                else:
                    error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(18, col_index)))
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
        compulsory_fields = [2, 4, 6]

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, col_index), 'NA', columns[col_index] + ' is a Compulsory Field')

    def build_out_duplicate_key(cell_row, cell_col, new_mod):
        concat_list = []
        for row in range(19, master_files['xl_sheet_main'].nrows):
            concat_list.append((master_files['xl_sheet_main'].cell_value(row, 0), master_files['xl_sheet_main'].cell_value(row, 2) + master_files['xl_sheet_main'].cell_value(row, 4)))

        matches = 0
        for concat_str in concat_list:
            # Check if part no same modifier
            if master_files['xl_sheet_main'].cell_value(cell_row, 0) == concat_str[0]:
                if PRIMARY_KEY_1 == concat_str[1]:
                    matches += 1

        if matches == 1:
            print ('Duplicate Key check --- Pass (Primary key is unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, matches, 'Primary key is unique in submitted master')
        elif matches >1:
            print ('Duplicate Key check --- Fail (Primary key is not unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, matches, 'Primary key is not unique in submitted master')

    def build_out_part_no(cell_row, cell_col, new_mod):
        part_and_customer_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)

        # Construct list of concatenated TTC Parts No. and Customer Contract
        comparison_list_1 = []
        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            comparison_list.append(str(selected['backup_0'].sheet_by_index(0).cell_value(row,3)) + str(selected['backup_0'].sheet_by_index(0).cell_value(row,5)))

        if part_and_customer_no in comparison_list:
            print('TTC Parts No. check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2), 'NA', 'P/N + Customer Contract No. found in Customer Contract Details (System)')
        else:
            comparison_list_2 = []
            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].sheet_by_index(0).nrows):
                    comparison_list.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].sheet_by_index(0).cell_value(row,3) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].sheet_by_index(0).cell_value(row,5))
            except KeyError:
                pass

            if part_and_customer_no in comparison_list_2:
                print ('TTC Parts No. check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2), 'NA', 'P/N + Customer Contract No. found in Customer Contract Details (Submitted)')
            else:
                print ('TTC Parts No. check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2), 'NA', 'P/N + Customer Contract No. not found in Customer Contract Details (System/Submitted)')

    # Check if build-out indicator is blank
    def build_out_indicator(cell_row, cell_col, new_mod):

        build_out_priority = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1))
        change_type = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2))
        build_out_date = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3))
        last_customer_order_month = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))
        stock_quantity = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+5))
        order_lot_flag = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+6))
        flag_2_parts = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+7))
        last_supplier_delivery_date = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+8))

        build_out_list = [build_out_priority, change_type, build_out_date, last_customer_order_month, stock_quantity, order_lot_flag, flag_2_parts, last_supplier_delivery_date]

        # If MOD build-out, check if mod to build-out indicator = N
        if new_mod == 'MOD':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N':
                if all(x == '' for x in build_out_list):
                    print ('Build-out Indicator --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join(build_out_list), 'Build-out fields left validated as blank for Build-out Indicator = N')

                else:
                    print ('Build-out Indicator --- Fail')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join(build_out_list), 'Build-out fields should be left blank if Build-out Indicator = N')

                return master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
            elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y':
                if all(x != '' for x in build_out_list):
                    print ('Build-out Indicator --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join(build_out_list), 'Build-out fields left validated as filled for Build-out Indicator = Y')

                else:
                    print ('Build-out Indicator --- Fail')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join(build_out_list), 'Build-out fields should be filled if Build-out Indicator = Y')
            else:
                print ('Build-out Indicator --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Fixed: Y or N')
        # If new build-out, check if indicator is Y
        else:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y':
                if all(x != '' for x in build_out_list):
                    print ('Build-out Indicator --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join(build_out_list), 'Build-out fields left validated as filled for Build-out Indicator = Y')

                else:
                    print ('Build-out Indicator --- Fail')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join(build_out_list), 'Build-out fields should be filled if Build-out Indicator = Y')
            elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N':
                print ('Build-out Indicator --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Build-out Indicator should not be N')
            else:
                print ('Build-out Indicator --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Fixed: Y or N')

        return master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

    # Check should compare against Build-out Priority Master
    def build_out_priority(cell_row, cell_col, new_mod):
        if any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == x for x in ('A', 'B', 'C', 'D', 'E')):
            print ('Build-out priority check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Build-out Priority found in Build-out Priority Master')
        else:
            print ('Build-out priority check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Build-out Priority not found in Build-out Priority Master')

    def build_out_change_type(cell_row, cell_col, new_mod):
        if any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == x for x in ('ECI', 'Localize', 'Runout', 'Others')):
            print ('Change Type check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Change Type Validated')
        else:
            print ('Change Type check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Change Type is case-sensitive and can only be: ECI, Localization, Runout or Others')

    def build_out_date(cell_row, cell_col, new_mod):
        # Convert strings into date format
        build_out_date = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        last_customer_order_month = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)
        last_supplier_delivery_date = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+5)

        reference_display = 'Last Customer Order Month: {}, Last Supplier Delivery Date: {}'.format(last_customer_order_month, last_supplier_delivery_date)

        try:
            date_1 = time.strptime(str(build_out_date),"%d %b %Y")
            date_2 = time.strptime(str(last_customer_order_month), "%b %Y")
            date_3 = time.strptime(str(last_supplier_delivery_date), "%d %b %Y")
        except ValueError: # Date doesn't make sense (e.g. 31 Jun 2016)
            print ('Build-out Date check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', build_out_date, reference_display, 'Dates Invalid')
            return

        # Last Customer Order Month should be before Last supplier Delivery Date which should be before Build-out Date
        if date_2 < date_3:
            if date_3 < date_1:
                print('Build-out Date check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', build_out_date, reference_display, 'Dates Validated')
            else:
                print('Build-out Date check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', build_out_date, reference_display, 'Last Supplier Delivery Date must be before Build-out Date')
        else:
            print('Build-out Date check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', build_out_date, reference_display, 'Last Customer Order Month must be before Last Supplier Delivery Date')

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
    def build_out_mod_reference(cell_row):

        # Get concat key
        part_no_customer_contract = str(master_files['xl_sheet_main'].cell_value(cell_row, 2)) + master_files['xl_sheet_main'].cell_value(cell_row, 4)

        # Extract all backup concat into list
        comparison_list_1 = []
        for row in range(19, selected['backup_16'].sheet_by_index(0).nrows):
            comparison_list_1.append((row, str(selected['backup_16'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_16'].sheet_by_index(0).cell_value(row, 4)))

        # Find backup row
        backup_row = 0
        for concat_str in comparison_list_1:
            if part_no_customer_contract == concat_str[1]:
                backup_row = concat_str[0]

        # If cannot find, return False
        if backup_row == 0:
            print ('MOD Reference check --- Fail')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Cannot find Backup Row in system for MOD part')
            return False

        if backup_row >= 0:
            # Extract contents of backup row
            backup_row_contents = []
            for col in range(0, selected['backup_16'].sheet_by_index(0).ncols):
                backup_row_contents.append(selected['backup_16'].sheet_by_index(0).cell_value(backup_row, col))

        submitted_contents = []
        validate_count = 0
        for col in range(0, 15): # Hard Code MAX COLUMN

            submitted_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            # ------- GET COLOUR INDEX ------- #
            xf_list = master_files['xl_workbook'].xf_list[master_files['xl_sheet_main'].cell_xf_index(cell_row, col)]
            cell_font = master_files['xl_workbook'].font_list[xf_list.font_index]
            # ------- GET COLOUR INDEX ------- #

            if all(col != x for x in (0, 1)):
                # IF RED
                if cell_font.colour_index == 10:
                    try:
                        if int(master_files['xl_sheet_main'].cell_value(cell_row, col)) != int(backup_row_contents[col]):
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
                            if int(master_files['xl_sheet_main'].cell_value(cell_row, col)) == int(backup_row_contents[col]):
                                validate_count += 1
                            else:
                                print ('MOD Reference Check --- Fail (%s BLACK but MOD)' % columns[col])
                                update_df('MOD', columns[col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, col), backup_row_contents[col], 'Field is indicated as \'UNCHANGED (BLACK)\' but different from system')
                        except ValueError:
                            print ('MOD Reference Check --- Fail (%s BLACK but MOD)' % columns[col])
                            update_df('MOD', columns[col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, col), backup_row_contents[col], 'Field is indicated as \'UNCHANGED (BLACK)\' but different from system')

        if validate_count == len(range(2, 15)): # Hard Code MAX COLUMN
            print ('MOD Reference check --- Pass')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', str(submitted_contents), str(backup_row_contents), 'Fields are correctly coloured to indicate \'TO CHANGE\'')

        return True

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
                if file.find('MRS_BuildOut') != -1:
                    selected['backup_16'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)

        if len(selected) == len(required):
            print('Successfully loaded all backup masters')
            print ('-' * 60)
            print ()
        else:
            # Allow partial retrieval for MOD parts, only if no NEW entries
            change_list = []
            for row in range(19, master_files['xl_sheet_main'].nrows):
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
        for row in range(19, master_files['xl_sheet_main'].nrows):

            # Set constant for primary keys per master check
            PRIMARY_KEY_1 = master_files['xl_sheet_main'].cell_value(row, 2)
            PRIMARY_KEY_2 = master_files['xl_sheet_main'].cell_value(row, 4)

            # If blank, skip row
            if master_files['xl_sheet_main'].cell_value(row, 2) == '':
                continue

            # Print Part no. header
            print ('%s: Part No. %s' % (str(master_files['xl_sheet_main'].cell_value(row, 0)), str(master_files['xl_sheet_main'].cell_value(row, 2))))
            print()

            # If float, prompt user to change format to string
            if isinstance(master_files['xl_sheet_main'].cell_value(row, 2), float):
                print ('Please change format of Row %d TTC Part No. to Text' % row)
                print ('-' * 10)
                update_df(master_files['xl_sheet_main'].cell_value(row, 0), columns[2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Please change format of Row ' + str(row + 1) + 'P/N to Text')
                continue

            # Conditional for NEW parts
            if str(master_files['xl_sheet_main'].cell_value(row, 0)) == 'NEW':
                check_maximum_length(row, 'NEW')
                check_compulsory_fields(row, 'NEW')
                build_out_duplicate_key(row, 2, 'NEW')
                build_out_part_no(row, 2, 'NEW')
                build_out_indicator(row, 6, 'NEW')
                build_out_priority(row, 7, 'NEW')
                build_out_change_type(row, 8, 'NEW')
                build_out_date(row, 9, 'NEW')
            # Conditional for MOD parts
            else:
                date_check_cycle = 0
                cols_to_check = get_mod_columns(row)
                if len(cols_to_check) != 0:
                    print('User wishes to MOD the following columns:')
                    for col in cols_to_check:
                        print('%s: %s' % (columns[col+2], master_files['xl_sheet_main'].cell_value(row, col+2)))
                    print()

                    if build_out_mod_reference(row):

                        check_maximum_length(row, 'MOD')
                        check_compulsory_fields(row, 'MOD')
                        build_out_duplicate_key(row, 2, 'MOD')

                        for col in cols_to_check:
                            # Mod: Module Group Code
                            if any(col+2 == x for x in (2, 4)):
                                print ('%s cannot be modded' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'Cannot be modded')
                            # Mod: Build-out Indicator, return indicator
                            if col+2 == 6:
                                build_out_indicator(row, 6, 'MOD')

                            # If build-out indicator N, no need to check the following fields.
                            # build_out_indicator already checks for blanks
                            indicator = master_files['xl_sheet_main'].cell_value(row, 6)

                            if indicator == 'N':
                                pass
                            elif indicator == 'Y':
                                # Mod: Build-out Priority
                                if col+2 == 7:
                                    build_out_priority(row, 7, 'MOD')
                                # Mod: Change Type
                                if col+2 == 8:
                                    build_out_change_type(row, 8, 'MOD')
                                # Mod: Build-out Date, Last Customer Order Month, Last Supplier Delivery Date
                                if any(col+2 == x for x in (9, 10, 14)):
                                    if date_check_cycle == 0:
                                        build_out_date(row, 9, 'MOD')
                                        date_check_cycle += 1
                                # Mod: optional columns
                                if (any(col+2 == x for x in (3, 5, 11, 12, 13))):
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
