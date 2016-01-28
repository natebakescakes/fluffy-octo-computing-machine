import time
import datetime
from os import listdir

import xlrd
import pandas as pd

from master_data import region_master
from master_data import west_export

# Supplier Parts - Open required workbooks and check against
def supplier_parts(master_files, path):
    # Dictionary of columns
    columns = {
        0: "NEW/MOD",
        1: "Reason for the change",
        2: "TTC Parts No.",
        3: "Supplier Code",
        4: "Supplier Parts No.",
        5: "Supplier Back No.",
        6: "Supplier Parts Name",
        7: "Exp Language Part Name",
        8: "Exp Country",
        9: "Current SRBQ",
        10: "Next SRBQ",
        11: "Next SRBQ Apply Date",
        12: "Big Parts",
        13: "Non Order Matrix",
        14: "Supplier Box Length",
        15: "Supplier Box Width",
        16: "Supplier Box Height",
        17: "Supplier Box M3",
        18: "Supplier Area 1",
        19: "Supplier Area 2"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        1: "Parts Master",
        2: "Customer Parts Master",
        3: "Supplier Parts Master"
    }

    def check_newmod_field(cell_row, cell_col):
        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != x for x in ('NEW', 'MOD')):
            print ('NEW/MOD check --- Fail')
            update_df(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Please check NEW/MOD field for whitespace')

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 20))

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
        compulsory_fields = [2, 3, 4, 6, 12, 13, 14, 15, 16, 17]

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', columns[col_index] + ' is a Compulsory Field')

    # Check for duplicate primary keys
    def supplier_parts_duplicate_key(cell_row, cell_col, new_mod):
        concat_list = []
        for row in range(9, master_files['xl_sheet_main'].nrows):
            concat_list.append((master_files['xl_sheet_main'].cell_value(row, 0), str(master_files['xl_sheet_main'].cell_value(row, 2)) + master_files['xl_sheet_main'].cell_value(row, 3)))

        matches = 0
        for concat_str in concat_list:
            # Check if part no same modifier
            if master_files['xl_sheet_main'].cell_value(cell_row, 0) == concat_str[0]:
                if PRIMARY_KEY_1 + PRIMARY_KEY_2 == concat_str[1]:
                    matches += 1

        if matches == 1:
            print ('Duplicate Key check --- Pass (Primary key is unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches, 'Primary key is unique in submitted master')
        elif matches >1:
            print ('Duplicate Key check --- Fail (Primary key is not unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches, 'Primary key is not unique in submitted master')

    # Match TTC Parts Number + Supplier Code in Customer Contract Details
    def supplier_parts_part_no_1(cell_row, cell_col, new_mod):
        part_and_supplier_code = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)

        comparison_list_1 = []
        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 11))

        matches_1 = 0
        for concat_str in comparison_list_1:
            if part_and_supplier_code == concat_str:
                matches_1 += 1

        if matches_1 == 1:
            print ('TTC Part No. check 1 --- Pass (System)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches_1, 'Registered in Customer Contract Details (System)')
        else:
            try:
                comparison_list_2 = []
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                    comparison_list_2.append(str(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3)) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 11))

                matches_2 = 0
                for concat_str in comparison_list_2:
                    if part_and_supplier_code == concat_str:
                        matches_2 += 1

                if matches_2 >= 1:
                    print ('TTC Part No. check 1 --- Pass (Submitted)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches_2, 'P/N + Supplier Code found in Customer Contract Details (Submitted)')
                else:
                    print ('TTC Part No. check 1 --- Fail (Cannot find part in System or Submitted)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches_2, 'P/N + Supplier Code not found in Customer Contract Details (Submitted/System)')
            except KeyError:
                print ('TTC Part No. check 1 --- Fail (Cannot find part in System)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, 'NA', 'P/N + Supplier Code not found in Customer Contract Details (System), no submitted Customer Contract Details')

    # TTC Parts No. + Supplier Code should not already exist in system
    def supplier_parts_part_no_2(cell_row, cell_col, new_mod):
        part_and_supplier_code = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)

        comparison_list_1 = []
        for row in range(9, selected['backup_3'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_3'].sheet_by_index(0).cell_value(row, 2) + selected['backup_3'].sheet_by_index(0).cell_value(row, 3))

        if part_and_supplier_code not in comparison_list_1:
            print ('TTC Part No. check 2 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1 + PRIMARY_KEY_2, 'NA', 'P/N + Supplier Code not already in system')
        else:
            backup_row_contents = []
            for row in range(9, selected['backup_3'].sheet_by_index(0).nrows):
                if part_and_supplier_code == str(selected['backup_3'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_3'].sheet_by_index(0).cell_value(row, 3):
                    for col in range(0, 20): # Hard Code column range
                        backup_row_contents.append(selected['backup_3'].sheet_by_index(0).cell_value(row, col))

            submitted_row_contents = []
            for col in range(0, 20): # Hard Code column range
                submitted_row_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            discrepancy_reference = []
            for i, cell_value in enumerate(submitted_row_contents):
                if cell_value != backup_row_contents[i] and all(i != x for x in (0, 1, 2, 3)):
                    discrepancy_reference.append(columns[i])

            if len(discrepancy_reference) == 0:
                discrepancy_reference.append('Submitted has no differences from system')

            print ('TTC Part No. check 2 --- Fail (Duplicate in system)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, ', '.join(discrepancy_reference), 'P/N + Supplier Code already registered in system')

    # Supplier Back No. must be input for 0TBSM, TBSJ, TBSQ, TBSR
    def supplier_parts_back_no(cell_row, cell_col, new_mod):
        required_suppliers = [
            'TH-0TBSM',
            'TH-TBSJ',
            'TH-TBSJ2',
            'TH-TBJ1',
            'TH-TBSQ',
            'TH-TBSR'
        ]

        if (any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2) == x for x in required_suppliers)):
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != '':
                print ('Supplier Back No. check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'Supplier Back No. inputted')
            else:
                print ('Supplier Back No. check --- Fail (Must be filled in)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'Supplier Back No. must be inputted for this Supplier Code')
        else:
            print ('Supplier Back No. check --- Pass (No need to fill)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'Supplier Back No. need not be filled in')

    # Must be in Region Master and match with Supplier Code
    def supplier_parts_exp_country(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in region_master:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5)[:2].find(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) != -1:
                print ('Exp Country check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5), 'In Region Master and matches with Supplier Code')
            elif (any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == x for x in ('C1', 'C2', 'C3')) and master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5)[:2] == 'CN') or (any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == x for x in ('I1', 'I2', 'I3')) and master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5)[:2] == 'IN'):
                print ('Exp Country check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5), 'In Region Master and matches with Supplier Code')
            else:
                print ('Exp Country check --- Fail (%s does not match with Supplier Code)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5), 'Does not match with Supplier Code')
        else:
            print ('Exp country check --- Fail (%s is not found in region master)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Not found in region master')

    def supplier_parts_next_srbq(cell_row, cell_col, new_mod):
        if new_mod == 'NEW':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != '' and master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) == '':
                print ('SRBQ check 1 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)) + ', ' + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)), 'NA', 'NEW Part current SRBQ and next SRBQ validated')
            else:
                print ('SRBQ check 1 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)) + ', ' + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)), 'NA', 'For NEW Parts, current SRBQ must be blank, next SRBQ must be filled')

        srbq = int(master_files['xl_sheet_main'].cell_value(cell_row, cell_col))

        part_no_supplier_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-8) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7)

        # Get list of P/N + Customer Code in CCD System
        part_no_customer_code_list = []
        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            if selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 11) == part_no_supplier_code:
                part_no_customer_code_list.append(selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 4))

        try:
            for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                if str(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3)) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 11) == part_no_supplier_code and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'NEW':
                    part_no_customer_code_list.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 4))
                elif additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'MOD'and str(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3)) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 11) == part_no_supplier_code:
                    if len(part_no_customer_code_list) == 1:
                        part_no_customer_code_list.remove(part_no_customer_code_list[0])
                    part_no_customer_code_list.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 4))

        except KeyError:
            pass

        if len(part_no_customer_code_list) > 1:
            print ('SRBQ check 2 --- Warning')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', srbq, str(part_no_customer_code_list), 'This part is used by multiple customers, please notify user especially if customers are of different Imp Country')
        elif len(part_no_customer_code_list) == 0:
            print ('SRBQ check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', srbq, 'NA', 'Cannot find P/N + Supplier Code in CCD')
            return

        # Get list of SPQ in Customer Parts
        spq_ol_list = []
        for concat_str in part_no_customer_code_list:
            for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
                if selected['backup_2'].sheet_by_index(0).cell_value(row, 2) + selected['backup_2'].sheet_by_index(0).cell_value(row, 3) == concat_str:
                    spq_ol_list.append((selected['backup_2'].sheet_by_index(0).cell_value(row, 2) + selected['backup_2'].sheet_by_index(0).cell_value(row, 3), selected['backup_2'].sheet_by_index(0).cell_value(row, 10), selected['backup_2'].sheet_by_index(0).cell_value(row, 12)))

        try:
            # Append NEW parts, MOD existing parts
            for concat_str in part_no_customer_code_list:
                for row in range(9, additional['TNM_CUSTOMER_PARTS_MASTER'].nrows):
                    if str(additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3) == concat_str and additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 0) == 'NEW':
                        spq_ol_list.append((str(additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 11), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 12)))
                    # If MOD, (concat, spq, ol) already in spq_ol_list, replace with MOD row.
                    elif str(additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3) == concat_str and additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 0) == 'MOD':
                        # Remove original tuple
                        for tuple in spq_ol_list:
                            if tuple[0] == additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2) + additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3):
                                spq_ol_list.remove(tuple)
                        # Replace with new tuple
                        if additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 11) != '':
                            spq_ol_list.append((str(additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 11), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 12)))
                        else:
                            spq_ol_list.append((str(additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 10), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 12)))

        except KeyError:
            pass

        if len(spq_ol_list) == 0:
            print ('SRBQ check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', srbq, str(part_no_customer_code_list), 'Cannot find P/N + Customer Code in Customer Parts Master')
            return

        # Make spq into list, then single value if all spq values are the same
        # Make ol into list, then single value if all ol values are the same
        spq_list, ol_list = [], []
        for tuple in spq_ol_list:
            spq_list.append(tuple[1])
            ol_list.append(tuple[2])
        spq_list = list(set(spq_list))
        ol_list = list(set(ol_list))

        if len(spq_list) == 1:
            spq = int(spq_list[0])
        else:
            print ('SRBQ check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', srbq, str(spq_list), 'Multiple SPQ values found, inform user that SPQ of corresponding customer parts do not match other SPQs that are registered in system')
            return

        if len(ol_list) == 1:
            orderlot = int(ol_list[0])
        else:
            print ('SRBQ check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', srbq, str(ol_list), 'Multiple OL values found, inform user that OL of corresponding customer parts do not match other OL that are registered in system')
            return

        print (srbq, spq)
        if srbq % spq == 0 or spq % srbq == 0:
            if orderlot % srbq == 0:
                print ('SRBQ check 2 --- Pass (OL: %d, SPQ: %d, SRBQ: %d)' % (orderlot, spq, srbq))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', srbq, 'OL: ' + str(orderlot) + ', SPQ: ' + str(spq), 'SPQ:SRBQ = 1:N or N:1, OL:SRBQ = N:1')
            else:
                print ('SRBQ check 2 --- Fail (OL: %d, SPQ: %d, SRBQ: %d)' % (orderlot, spq, srbq))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', srbq, 'OL: ' + str(orderlot) + ', SPQ: ' + str(spq), 'OL:SRBQ != N:1')
        else:
            print ('SRBQ check 2 --- Fail (OL: %d, SPQ: %d, SRBQ: %d)' % (orderlot, spq, srbq))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', srbq, 'OL: ' + str(orderlot) + ', SPQ: ' + str(spq), 'SPQ:SRBQ != 1:N or N:1')

    # Must be 1st of this month or 1st of next month
    # If 1st of this month, part must not have any (non-cancelled) order in past or current month.
    # If Next_SRBQ is input, Next_SRBQ Apply Date must be input. If not, both must be blank.
    def supplier_parts_srbq_apply_date(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) != '':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != '':
                print ('SRBQ Apply Date check 1 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Both Next SRBQ and Apply Date are validated as filled')
            else:
                print ('SRBQ Apply Date check 1 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'If Next SRBQ is input, Apply Date must be input')
        else:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == '':
                print ('SRBQ Apply Date check 1 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Both Next SRBQ and Apply Date are validated as not filled')
            else:
                print ('SRBQ Apply Date check 1 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'If either Next SRBQ or Apply Date are blank, both must be blank')

        # Must be 1st of this month or 1st of next month
        # Extract Date
        try:
            apply_date = time.strptime(str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)),"%d %b %Y")
        except ValueError:
            print ('IP Specs Apply Date check --- Fail (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Check Date Format')

        # transform today into 1st of next month
        def add_month(current_date):
            year = int(current_date.year + current_date.month / 12)
            month = current_date.month % 12 + 1
            day = 1
            return datetime.date(year, month, day)

        # transform today into 1st of this month
        def this_month(current_date):
            year = current_date.year
            month = current_date.month
            day = 1
            return datetime.date(year, month, day)

        correct_apply_date = add_month(datetime.date.today()).timetuple()
        additional_order_apply_date = this_month(datetime.date.today()).timetuple()

        if 'apply_date' in locals():
            if apply_date == correct_apply_date:
                print ('SRBQ Apply Date check 2 --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', correct_apply_date), 'Apply date is 1st of future month')
            elif apply_date == additional_order_apply_date and master_files['xl_sheet_main'].cell_value(cell_row, 0) == 'MOD':
                print ('SRBQ Apply Date check 2 --- Warning')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', additional_order_apply_date), 'Apply date is 1st of this month, please check if part has non-cancelled order in past or current month')
            elif apply_date == additional_order_apply_date:
                print ('SRBQ Apply Date check 2 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', additional_order_apply_date), 'Apply date is 1st of this month')
            else:
                print ('SRBQ Apply Date check 2 --- Fail (must be 1st of next month)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(time.strftime('%d %b %Y', correct_apply_date)) + ' or ' + str(time.strftime('%d %b %Y', correct_apply_date)), 'Must be 1st of Next Month or Current Month')

    # Big Parts & Non-order Matrix cannot both have 'Y' flag
    def supplier_parts_big_parts(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y' and master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1) == 'Y':
            print ('Big Parts & Non Order Matrix check --- Fail (Cannot both \'Y\')')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'Big Parts: ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + ', Non Order Matrix: ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 'NA', "Big Part and Non Order Matrix Fields cannot both be 'Y'")
        else:
            print ('Big Parts & Non Order Matrix check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'Big Parts: ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + ', Non Order Matrix: ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 'NA', "Big Part and Non Order Matrix not both 'Y'")

    # Supplier Box Specs cannot be zero
    def supplier_parts_box_specs(cell_row, cell_col, new_mod):
        box_length = round(float(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)), 3)
        box_width = round(float(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), 3)
        box_height = round(float(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), 3)

        if box_length == 0 or box_width == 0 or box_height == 0:
            print ('Supplier Box Specs check --- Fail (LWH cannot be 0)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'L: ' + str(box_length) + ', W: ' + str(box_width) + ', H: ' + str(box_height), 'NA', "LWH cannot be 0")
        else:
            print ('Supplier Box Specs check --- Pass (%.3f * %.3f * %.3f)' % (box_length, box_width, box_height))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'L: ' + str(box_length) + ', W: ' + str(box_width) + ', H: ' + str(box_height), 'NA', "LWH not 0")

    # Validate calculated value for Box M3
    # =ROUND(ROUND(M11,3)*ROUND(N11,3)*ROUND(O11,3),3)
    def supplier_parts_box_m3(cell_row, cell_col, new_mod):

        # Transcribe excel formula to python
        correct_value = round(
                        round(float(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3)), 3) *
                        round(float(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)), 3) *
                        round(float(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)), 3), 3)

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == correct_value:
            print ("Supplier Box Specs check --- Pass (Formula Validated)")
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(correct_value), 'Formula Validated')
        else:
            print ('Supplier Box Specs check --- Fail (M3 should be %f)' % correct_value)
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(correct_value), 'Box M3 not calculated correctly')

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
    def supplier_parts_mod_reference(cell_row):

        # Get concat key
        part_no_supplier_code = str(master_files['xl_sheet_main'].cell_value(cell_row, 2)) + master_files['xl_sheet_main'].cell_value(cell_row, 3)

        # Extract all backup concat into list
        comparison_list_1 = []
        for row in range(9, selected['backup_3'].sheet_by_index(0).nrows):
            comparison_list_1.append((row, str(selected['backup_3'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_3'].sheet_by_index(0).cell_value(row, 3)))

        # Find backup row
        backup_row = 0
        for concat_str in comparison_list_1:
            if part_no_supplier_code == concat_str[1]:
                backup_row = concat_str[0]

        # If cannot find, return False
        if backup_row == 0:
            print ('MOD Reference check --- Fail')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Cannot find Backup Row in system for MOD part')
            return False

        if backup_row >= 0:
            # Extract contents of backup row
            backup_row_contents = []
            for col in range(0, selected['backup_3'].sheet_by_index(0).ncols):
                backup_row_contents.append(selected['backup_3'].sheet_by_index(0).cell_value(backup_row, col))

        submitted_contents = []
        validate_count = 0
        for col in range(0, 18): # Hard Code MAX COLUMN

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

        if validate_count == len(range(2, 18)): # Hard Code MAX COLUMN
            print ('MOD Reference check --- Pass')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', str(submitted_contents), str(backup_row_contents), 'Fields are correctly coloured to indicate \'TO CHANGE\'')

        return True

    # Check if company code is in WEST
    def supplier_parts_west_part_master(cell_row, cell_col, new_mod):
        # Get concat key
        part_no = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col))

        # Find WEST Fields in system
        backup_exp_west_fields = []
        for row in range(9, selected['backup_1'].sheet_by_index(0).nrows):
            if selected['backup_1'].sheet_by_index(0).cell_value(row, 2) == part_no:
                backup_exp_west_fields.append((selected['backup_1'].sheet_by_index(0).cell_value(row, 8), selected['backup_1'].sheet_by_index(0).cell_value(row, 9), selected['backup_1'].sheet_by_index(0).cell_value(row, 10), selected['backup_1'].sheet_by_index(0).cell_value(row, 11), selected['backup_1'].sheet_by_index(0).cell_value(row, 12)))

        try:
            for row in range(9, additional['TNM_PARTS_MASTER'].nrows):
                if additional['TNM_PARTS_MASTER'].cell_value(row, 2) == part_no:
                    backup_exp_west_fields.append((additional['TNM_PARTS_MASTER'].cell_value(row, 8), additional['TNM_PARTS_MASTER'].cell_value(row, 9), additional['TNM_PARTS_MASTER'].cell_value(row, 10), additional['TNM_PARTS_MASTER'].cell_value(row, 11), additional['TNM_PARTS_MASTER'].cell_value(row, 12)))

        except KeyError:
            pass

        company_code_list = []
        for tuple in backup_exp_west_fields:
            company_code_list.append(tuple[0]) # Check company code only for now

        company_code_list = list(set(company_code_list)) # Remove duplicates

        try:
            export_office_code = west_export[master_files['xl_sheet_main'].cell_value(cell_row, cell_col+6)]['Office Code']
        except KeyError:
            print ('Parts Master WEST check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_no, 'NA', 'Exp Country does not use WEST')
            return

        if export_office_code in company_code_list:
            print ('Parts Master WEST check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_no + ', ' + export_office_code, company_code_list, 'Company Code registered in Parts Master')
        else:
            if export_office_code == 'S556':
                print ('Parts Master WEST check --- Warning')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', part_no + ', ' + export_office_code, company_code_list, 'Company Code not registered in Parts Master (Optional for TW)')
            else:
                print ('Parts Master WEST check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', part_no + ', ' + export_office_code, company_code_list, 'Company Code not registered in Parts Master, please add MOD row in Parts Master to include Company Code')

    def supplier_parts_discontinued(cell_row, cell_col):
        part_no_supplier_code = PRIMARY_KEY_1 + PRIMARY_KEY_2

        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            if part_no_supplier_code == str(selected['backup_0'].sheet_by_index(0).cell_value(row, 3)) + selected['backup_0'].sheet_by_index(0).cell_value(row, 11):
                if selected['backup_0'].sheet_by_index(0).cell_value(row, 8) == 'N':
                    print ('Discontinued check --- Pass')
                    update_df('MOD', columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_no_supplier_code, selected['backup_0'].sheet_by_index(0).cell_value(row, 8), 'Part has not been discontinued')
                else:
                    print ('Discontinued check --- Fail')
                    update_df('MOD', columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', part_no_supplier_code, selected['backup_0'].sheet_by_index(0).cell_value(row, 8), 'Part has already been discontinued')


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
                if file.find('MRS_PartMaster') != -1:
                    selected['backup_1'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_CustomerPartsMaster') != -1:
                    selected['backup_2'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_SupplierPartsMaster') != -1:
                    selected['backup_3'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
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

        # Checkpoints
        for row in range(9, master_files['xl_sheet_main'].nrows):

            # Set constant for primary keys per master check
            PRIMARY_KEY_1 = master_files['xl_sheet_main'].cell_value(row, 2)
            PRIMARY_KEY_2 = master_files['xl_sheet_main'].cell_value(row, 3)

            # If blank, skip row
            if master_files['xl_sheet_main'].cell_value(row, 2) == '':
                continue

            # Print Part no. header
            print ('%s: Part No. %s' % (str(master_files['xl_sheet_main'].cell_value(row, 0)), str(master_files['xl_sheet_main'].cell_value(row, 2))))
            print()

            # If float, prompt user to change format to String
            if isinstance(master_files['xl_sheet_main'].cell_value(row, 2), float):
                print ('Please change format of Row %d TTC Part No. to Text' % row)
                print ('-' * 10)
                update_df(master_files['xl_sheet_main'].cell_value(row, 0), columns[2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Please change format of Row ' + str(row + 1) + ' P/N to Text')
                continue

            check_newmod_field(row, 0)

            # Conditional for NEW parts
            if str(master_files['xl_sheet_main'].cell_value(row, 0)).strip(' ') == 'NEW':
                check_maximum_length(row, 'NEW')
                check_compulsory_fields(row, 'NEW')
                supplier_parts_duplicate_key(row, 2, 'NEW')
                supplier_parts_part_no_1(row, 2, 'NEW')
                supplier_parts_part_no_2(row, 2, 'NEW')
                supplier_parts_west_part_master(row, 2, 'NEW')
                supplier_parts_back_no(row, 5, 'NEW')
                supplier_parts_exp_country(row, 8, 'NEW')
                supplier_parts_next_srbq(row, 10, 'NEW')
                supplier_parts_srbq_apply_date(row, 11, 'NEW')
                supplier_parts_big_parts(row, 12, 'NEW')
                supplier_parts_box_specs(row, 14, 'NEW')
                supplier_parts_box_m3(row, 17, 'NEW')
            # Conditional for MOD parts
            else:
                cols_to_check = get_mod_columns(row)
                if len(cols_to_check) != 0:
                    print('User wishes to MOD the following columns:')
                    for col in cols_to_check:
                        print('%s:' % columns[col+2])
                    print()

                    if supplier_parts_mod_reference(row):
                        check_maximum_length(row, 'MOD')
                        check_compulsory_fields(row, 'MOD')
                        supplier_parts_duplicate_key(row, 2, 'MOD')
                        supplier_parts_discontinued(row, 2)

                        for col in cols_to_check:
                            # Mod: TTC Parts No., Supplier Code
                            if (any(col+2 == x for x in (2, 3, 9))):
                                print ('%s cannot be modded' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'Cannot be modded')
                            # Mod: Supplier Back No.
                            if col+2 == 5:
                                supplier_parts_back_no(row, col+2, 'MOD')
                            # Mod: Exp Country
                            if col+2 == 8:
                                supplier_parts_exp_country(row, col+2, 'MOD')
                            # Mod: Next_SRBQ
                            if col+2 == 10:
                                supplier_parts_next_srbq(row, col+2, 'MOD')
                            # Mod: Next_SRBQ Apply Date
                            if col+2 == 11:
                                supplier_parts_srbq_apply_date(row, col+2, 'MOD')
                            # Mod: Big Parts, Non Order Matrix
                            if (any(col+2 == x for x in (12, 13))):
                                supplier_parts_big_parts(row, 12, 'MOD')
                            # Mod: Supplier Box LWH
                            if (any(col+2 == x for x in (14, 15, 16))):
                                supplier_parts_box_specs(row, 14, 'MOD')
                            # Mod: Supplier Box M3
                            if col+2 == 17:
                                supplier_parts_box_m3(row, col+2, 'MOD')
                            # Mod: optional columns
                            if (any(col+2 == x for x in (4, 6, 7, 18, 19))):
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
