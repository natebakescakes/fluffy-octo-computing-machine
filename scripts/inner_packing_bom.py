import time
from os import listdir

import xlrd
import pandas as pd

# Customer Parts - Open required workbooks and check against
def inner_packing_bom(master_files, path):
# Dictionary of columns
    columns = {
        0: "NEW",
        1: "Exp WH Code",
        2: "TTC Parts No.",
        3: "Customer Code",
        4: "SPQ",
        5: "Material No.",
        6: "Sequence No.",
        7: "Material Quantity"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        1: "Customer Parts Master",
        2: "Inner Packing BOM",
    }

    def check_newmod_field(cell_row, cell_col):
        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != x for x in ('NEW', 'MOD')):
            print ('NEW/MOD check --- Fail')
            update_df(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Please check NEW/MOD field for whitespace')

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 8))

        validate_count = 0
        error_fields = []

        # validate if num   ber
        for col_index in working_columns:
            # validate if blank: Fixed (Probably has its own check, unless Y/N)
            if master_files['xl_sheet_main'].cell_value(8, col_index) == '':
                validate_count += 1
                continue

            try:
                # conditional if number
                maximum_length = int(master_files['xl_sheet_main'].cell_value(8, col_index))
                # check if submitted is integer, remove '.0'
                try:
                    if len(str(int(master_files['xl_sheet_main'].cell_value(cell_row, col_index)))) <= maximum_length:
                        validate_count += 1
                        continue
                    else:
                        error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(8, col_index)))
                        continue
                except ValueError:
                    if len(str(master_files['xl_sheet_main'].cell_value(cell_row, col_index))) <= maximum_length:
                        validate_count += 1
                        continue
                    else:
                        error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(8, col_index)))
                        continue
            except ValueError:
                # conditional if date
                if str(master_files['xl_sheet_main'].cell_value(8, col_index)) == 'dd mmm yyyy':
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
                                error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(8, col_index)))
                                continue
                        else:
                            error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(8, col_index)))
                            continue
                # conditional if Month
                elif str(master_files['xl_sheet_main'].cell_value(8, col_index)) == 'mmm yyyy':
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
                                error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(8, col_index)))
                                continue
                        else:
                            error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(8, col_index)))
                            continue
                # conditional if integer,integer
                # Only checks total length, not decimal places
                elif master_files['xl_sheet_main'].cell_value(8, col_index).find(',') != -1:
                    char_limit = master_files['xl_sheet_main'].cell_value(8, col_index).split(',')
                    try:
                        if len(str(master_files['xl_sheet_main'].cell_value(cell_row, col_index))) <= (int(char_limit[0]) + int(char_limit[1]) + 1):
                            validate_count += 1
                            continue
                        else:
                            error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(8, col_index)))
                            continue
                    except ValueError:
                        error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(8, col_index)))
                        continue
                else:
                    error_fields.append((col_index, master_files['xl_sheet_main'].cell_value(8, col_index)))
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
        compulsory_fields = list(range(1, 8))

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', columns[col_index] + ' is a Compulsory Field')

    # Check for duplicate primary keys
    def inner_packing_bom_duplicate_key(cell_row, cell_col, new_mod):
        concat_list = []
        for row in range(9, master_files['xl_sheet_main'].nrows):
            concat_list.append((master_files['xl_sheet_main'].cell_value(row, 0), master_files['xl_sheet_main'].cell_value(row, 1) + str(master_files['xl_sheet_main'].cell_value(row, 2)) + master_files['xl_sheet_main'].cell_value(row, 3) + str(master_files['xl_sheet_main'].cell_value(row, 4)) + master_files['xl_sheet_main'].cell_value(row, 5)))

        primary_key = master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4)

        matches = 0
        for concat_str in concat_list:
            # Check if part no same modifier
            if master_files['xl_sheet_main'].cell_value(cell_row, 0) == concat_str[0]:
                if primary_key == concat_str[1]:
                    matches += 1

        if matches == 1:
            print ('Duplicate Key check --- Pass (Primary key is unique in submitted master)')
            update_df(new_mod, 'Primary key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', primary_key, matches, 'Primary keys are unique in submitted master')
        elif matches >1:
            print ('Duplicate Key check --- Fail (Primary key is not unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', primary_key, matches, 'Primary keys are not unique in submitted master')

    # Match TTC Parts Number + Customer Code in Customer Contract Details for correct exp country
    def inner_packing_bom_part_no(cell_row, cell_col, new_mod):

        part_and_customer_code = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)

        full_system_submitted = []

        try:
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                full_system_submitted.append((selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 4), selected['backup_0'].sheet_by_index(0).cell_value(row, 10)))

            for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                full_system_submitted.append((str(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3)) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 4), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 10)))

        except KeyError:
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                full_system_submitted.append((selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 4), selected['backup_0'].sheet_by_index(0).cell_value(row, 10)))

        matches = []
        for tuple in full_system_submitted:
            if part_and_customer_code == tuple[0]:
                matches.append(tuple[1][:2])

        match_exp_code = False
        for exp_country in matches:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)[:2] == exp_country:
                match_exp_code = True

        if len(matches) > 0:
            if match_exp_code:
                print ('TTC Part No. check --- Pass (%s, %s)' % (part_and_customer_code, master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)))
                update_df(new_mod, columns[2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_and_customer_code, 'NA', 'P/N + Customer Code must be found in CCD for correct Exp Country')
            else:
                print ('TTC Part No. check --- Fail (%s does not match)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1))
                update_df(new_mod, columns[2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), matches, 'Exp Country does not match CCD')
        else:
            print ('TTC Part No. check --- Fail (%s cannot be found)' % part_and_customer_code)
            update_df(new_mod, columns[2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', part_and_customer_code, 'NA', 'Cannot find P/N + Customer Code in CCD')

    # SPQ must match TTC Parts No and Customer in Customer Parts master
    def inner_packing_bom_spq(cell_row, cell_col, new_mod):
        spq_inner_packing = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        part_customer_spq = {}
        try:
            for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
                part_customer_spq[selected['backup_2'].sheet_by_index(0).cell_value(row, 2) + selected['backup_2'].sheet_by_index(0).cell_value(row, 3)] = selected['backup_2'].sheet_by_index(0).cell_value(row, 10)

            for row in range(9, additional['TNM_CUSTOMER_PARTS_MASTER'].nrows):
                if additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 11) != '':
                    part_customer_spq[str(additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3)] = additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 11)
                else:
                    part_customer_spq[str(additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3)] = additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 10)

        except KeyError:
            for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
                part_customer_spq[str(selected['backup_2'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_2'].sheet_by_index(0).cell_value(row, 3)] = selected['backup_2'].sheet_by_index(0).cell_value(row, 10)

        try:
            if int(part_customer_spq[str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)]) == int(spq_inner_packing):
                print ('SPQ Check --- Pass (IPB: %s, CP: %s)' % (spq_inner_packing, part_customer_spq[str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)]))
                update_df(new_mod, columns[4], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', spq_inner_packing, part_customer_spq[str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)], 'SPQ matches P/N + Customer Code in CP')
            else:
                print ('SPQ Check --- Pass (IPB: %s, CP: %s)' % (spq_inner_packing, part_customer_spq[str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)]))
                update_df(new_mod, columns[4], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', spq_inner_packing, part_customer_spq[str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)], 'SPQ does not match P/N + Customer Code in CP')
        except KeyError:
            print ('SPQ Check --- Fail (Cannot find P/N + Customer Code in CP)')
            update_df(new_mod, columns[4], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', spq_inner_packing, 'NA', 'SPQ does not match P/N + Customer Code in CP')

    # Check if Material No. already registered in IPB Backup
    def inner_packing_bom_material(cell_row, cell_col, new_mod):
        material_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        backup_material_list = []
        for row in range(9, selected['backup_9'].sheet_by_index(0).nrows):
            backup_material_list.append(selected['backup_9'].sheet_by_index(0).cell_value(row, 5))

        if material_no in backup_material_list:
            print ('Material check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', material_no, 'NA', 'Material No. found already registered in Inner Packing BOM')
        else:
            print ('Material check --- Warning')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', material_no, 'NA', 'Material No. is new, please check System Screen whether it has already been registered correctly')

    # TTC Parts No + Material No (+ SPQ + Customer + Exp WH) must be unique
    # TTC Parts No + Sequence No (+ SPQ + Customer + Exp WH) must be unique
    def inner_packing_bom_concat(cell_row, cell_col, new_mod):
        concat_1 = master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2) + str(int(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3))) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4)

        concat_2 = master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2) + str(int(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3))) + str(int(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+5)))

        matches_1, matches_2 = 0, 0
        for concat_str in backup_concat_1:
            if concat_str == concat_1:
                matches_1 += 1

        for concat_str in backup_concat_2:
            if concat_str == concat_2:
                matches_2 += 1

        if matches_1 == 0:
            print ('Material No. concat check --- Pass (TTC Parts No + Material No (+ SPQ + Customer + Exp WH) is unique)')
            update_df(new_mod, columns[5], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', concat_1, 'NA', 'P/N + Material No. (+ SPQ + Customer + Exp WH) is unique)')
        else:
            print ('Material No. concat check --- Fail (TTC Parts No + Material No (+ SPQ + Customer + Exp WH) not unique)')
            update_df(new_mod, columns[5], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', concat_1, 'NA', 'P/N + Material No. (+ SPQ + Customer + Exp WH) is not unique)')

        if matches_2 == 0:
            print ('Sequence No. concat check --- Pass (TTC Parts No + Sequence No (+ SPQ + Customer + Exp WH) is unique)')
            update_df(new_mod, columns[6], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', concat_2, 'NA', 'P/N + Sequence No. (+ SPQ + Customer + Exp WH) is unique)')
        else:
            print ('Sequence No. concat check --- Fail (TTC Parts No + Sequence No (+ SPQ + Customer + Exp WH) not unique)')
            update_df(new_mod, columns[6], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', concat_2, 'NA', 'P/N + Sequence No. (+ SPQ + Customer + Exp WH) is not unique)')

    # Sequence no of inner packing materials must be in Sequence
    def inner_packing_bom_sequence(cell_row, cell_col, new_mod):
        concat_4_fields = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-4)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2))

        part_no_sequence_no = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-4)) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3)) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col))

        concat_sequence, part_no_sequence_no_list = [], []
        for row in range(9, master_files['xl_sheet_main'].nrows):
            concat_sequence.append((master_files['xl_sheet_main'].cell_value(row, cell_col-5) + str(master_files['xl_sheet_main'].cell_value(row, cell_col-4)) + master_files['xl_sheet_main'].cell_value(row, cell_col-3) + str(master_files['xl_sheet_main'].cell_value(row, cell_col-2)), str(master_files['xl_sheet_main'].cell_value(row, cell_col))))

            part_no_sequence_no_list.append((str(master_files['xl_sheet_main'].cell_value(row, cell_col-4)) + str(master_files['xl_sheet_main'].cell_value(row, cell_col-3)) + str(master_files['xl_sheet_main'].cell_value(row, cell_col-2)) + str(master_files['xl_sheet_main'].cell_value(row, cell_col))))

        sequence_list = []
        for concat_str in concat_sequence:
            if concat_4_fields == concat_str[0]:
                sequence_list.append(concat_str[1])

        correct_sequence = list(range(1, len(sequence_list)+1))

        match_count = 0
        for tuple in part_no_sequence_no_list:
            if part_no_sequence_no == tuple:
                match_count += 1

        if sequence_list.sort() == correct_sequence.sort():
            if match_count == 1:
                print ('Sequence No. check --- Pass (In sequence)')
                update_df(new_mod, columns[6], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), correct_sequence, 'Sequence No. of Materials is in correct sequence')
            else:
                print ('Sequence No. check --- Fail (Duplicate sequence no.)')
                update_df(new_mod, columns[6], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), match_count, 'Duplicate sequence no. for WH + P/N + Customer Code')
        else:
            print ('Sequence No. check --- Fail (Not in sequence)')
            update_df(new_mod, columns[6], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), correct_sequence, 'Sequence No. of Materials not in correct sequence')

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
                if file.find('MRS_CustomerPartsMaster') != -1:
                    selected['backup_2'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_InnerPackingBOM') != -1:
                    selected['backup_9'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)

        if len(selected) == len(required):
            print('Successfully loaded all backup masters')
            print ('-' * 60)
            print ()
        else:
            # Allow partial retrieval for MOD parts, only if no NEW entries
            change_list = []
            for row in range(9, master_files['xl_sheet_main'].nrows):
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

        # Concatenation check
        backup_concat_1, backup_concat_2 = [], []
        for row in range(9, selected['backup_9'].sheet_by_index(0).nrows):
            backup_concat_1.append(selected['backup_9'].sheet_by_index(0).cell_value(row, 1) +  str(selected['backup_9'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_9'].sheet_by_index(0).cell_value(row, 3) + str(int(selected['backup_9'].sheet_by_index(0).cell_value(row, 4))) + selected['backup_9'].sheet_by_index(0).cell_value(row, 5))

        for row in range(9, selected['backup_9'].sheet_by_index(0).nrows):
            backup_concat_1.append(selected['backup_9'].sheet_by_index(0).cell_value(row, 1) + str(selected['backup_9'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_9'].sheet_by_index(0).cell_value(row, 3) + str(selected['backup_9'].sheet_by_index(0).cell_value(row, 4)) + str(selected['backup_9'].sheet_by_index(0).cell_value(row, 6)))

        # Checkpoints
        material_type_list = []
        for row in range(9, master_files['xl_sheet_main'].nrows):

            # Set constant for primary keys per master check
            PRIMARY_KEY_1 = master_files['xl_sheet_main'].cell_value(row, 2)
            PRIMARY_KEY_2 = master_files['xl_sheet_main'].cell_value(row, 3)

            # Append Material Type to list
            material_type_list.append(master_files['xl_sheet_main'].cell_value(row, 5))

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

            check_newmod_field(row, 0)

            # Conditional for NEW parts
            if str(master_files['xl_sheet_main'].cell_value(row, 0)).strip(' ') == 'NEW':
                check_maximum_length(row, 'NEW')
                check_compulsory_fields(row, 'NEW')
                inner_packing_bom_duplicate_key(row, 1, 'NEW')
                inner_packing_bom_part_no(row, 2, 'NEW')
                inner_packing_bom_spq(row, 4, 'NEW')
                inner_packing_bom_material(row, 5, 'NEW')
                inner_packing_bom_concat(row, 1, 'NEW')
                inner_packing_bom_sequence(row, 6, 'NEW')
            # Conditional for MOD parts
            else:
                print ('Inner Packing BOM cannot be modded!')
                update_df(master_files['xl_sheet_main'].cell_value(row, 0), columns[2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, 0), 'NA', 'Inner Packing BOM cannot be modded')

            print ('-' * 10)

        # Pandas export to excel
        df = pd.DataFrame(check_dict)
        return df
