import time
import datetime
from os import listdir

import xlrd
import pandas as pd

from region_master import region_master
from west_data import west_import

# Customer Parts - Open required workbooks and check against
def customer_parts(master_files, path):
    # Dictionary of columns
    columns = {
        0: "NEW/MOD",
        1: "Reason for the change",
        2: "TTC Parts No.",
        3: "Customer Code",
        4: "Customer Parts No.",
        5: "Customer Parts Name",
        6: "Imp HS Code",
        7: "Imp Country Code",
        8: "Customer Back No.",
        9: "WEST Invoice Customer Parts No.",
        10: "Current SPQ",
        11: "Next SPQ",
        12: "Orderlot",
        13: "Orderlot Apply Date",
        14: "IP Size (L)",
        15: "IP Size (W)",
        16: "IP Size (H)",
        17: "IP Gross Weight",
        18: "IP Specs Apply Date",
        19: "Paired Parts",
        20: "TTC Parts No. of Paired Parts",
        21: "Paired Parts Flag",
        22: "Exp Logistics Back No.",
        23: "Inner Packing Time",
        24: "Fixed Location Code",
        25: "Pick Rack Location Code",
        26: "Heavy Parts"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        1: "Parts Master",
        2: "Customer Parts Master",
        3: "Supplier Parts Master",
        4: "Inner Packing BOM"
    }

    def check_newmod_field(cell_row, cell_col):
        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != x for x in ('NEW', 'MOD')):
            print ('NEW/MOD check --- Fail')
            update_df(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Please check NEW/MOD field for whitespace')

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 27))

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
                        if col_index == 18: # IP Specs Apply Date, can be <MRS UPLOAD DATE>
                            if master_files['xl_sheet_main'].cell_value(cell_row, col_index).upper() == '<MRS UPLOAD DATE>':
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
        compulsory_fields = [2, 3, 4, 5, 6, 12, 13, 14, 15, 16, 17, 18, 19, 22, 23, 24, 25, 26]

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', columns[col_index] + ' is a Compulsory Field')

    # Check for duplicate primary keys
    def customer_parts_duplicate_key(cell_row, cell_col, new_mod):
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

    # Match TTC Parts Number + Customer Code in Customer Contract Details
    def customer_parts_part_no_1(cell_row, cell_col, new_mod):
        part_and_customer_code = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)

        comparison_list_1 = []
        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 4))

        matches_1 = 0
        for concat_str in comparison_list_1:
            if part_and_customer_code == concat_str:
                matches_1 += 1

        if matches_1 >= 1:
            print ('TTC Part No. check 1 --- Pass (System)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches_1, 'Registered in Customer Contract Details (System)')
        else:
            try:
                comparison_list_2 = []
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                    comparison_list_2.append(str(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3)) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 4))

                matches_2 = 0
                for concat_str in comparison_list_2:
                    if part_and_customer_code == concat_str:
                        matches_2 += 1

                if matches_2 == 1:
                    print ('TTC Part No. check 1 --- Pass (Submitted)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches_2, 'P/N + Customer Code found in Customer Contract Details (Submitted)')
                elif matches_2 > 1:
                    print ('TTC Part No. check 1 --- Fail (Duplicate part)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches_2, 'Duplicate P/N + Customer Code in Customer Contract Details (Submitted)')
                else:
                    print ('TTC Part No. check 1 --- Fail (Cannot find part in System or Submitted)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches_2, 'P/N + Customer Code not found in Customer Contract Details (Submitted/System)')
            except KeyError:
                print ('TTC Part No. check 1 --- Fail (Cannot find part in System)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, 'NA', 'P/N + Customer Code not found in Customer Contract Details (System), no submitted Customer Contract Details')

    # TTC Parts No. + Customer Code should not already exist in system
    def customer_parts_part_no_2(cell_row, cell_col, new_mod):
        part_and_customer_code = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)

        comparison_list_1 = []
        for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_2'].sheet_by_index(0).cell_value(row, 2) + selected['backup_2'].sheet_by_index(0).cell_value(row, 3))

        if part_and_customer_code not in comparison_list_1:
            print ('TTC Part No. check 2 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1 + PRIMARY_KEY_2, 'NA', 'P/N + Customer Code not already in system')
        else:
            backup_row_contents = []
            for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
                if part_and_customer_no == str(selected['backup_2'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_2'].sheet_by_index(0).cell_value(row, 3):
                    for col in range(0, 27): # Hard Code column range
                        backup_row_contents.append(selected['backup_2'].sheet_by_index(0).cell_value(row, col))

            submitted_row_contents = []
            for col in range(0, 27): # Hard Code column range
                submitted_row_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            discrepancy_reference = []
            for i, cell_value in enumerate(submitted_row_contents):
                if cell_value != backup_row_contents[i] and all(i != x for x in (0, 1, 2, 3)):
                    discrepancy_reference.append(columns[i])

            if len(discrepancy_reference) == 0:
                discrepancy_reference.append('Submitted has no differences from system')

            print ('TTC Part No. check 2 --- Fail (Duplicate in system)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, ', '.join(discrepancy_reference), 'P/N + Customer Code already registered in system')

    # Part No. + Customer Code should be in Inner Packing BOM
    def customer_parts_part_no_3(cell_row, cell_col, new_mod):
        part_and_customer_code = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)

        comparison_list_1 = []
        try:
            for row in range(9, selected['backup_9'].sheet_by_index(0).nrows):
                comparison_list_1.append(selected['backup_9'].sheet_by_index(0).cell_value(row, 2) + selected['backup_9'].sheet_by_index(0).cell_value(row, 3))

            for row in range(9, additional['TNM_INNER_PACKING_BOM'].nrows):
                comparison_list_1.append(str(additional['TNM_INNER_PACKING_BOM'].cell_value(row, 2)) + additional['TNM_INNER_PACKING_BOM'].cell_value(row, 3))

        except KeyError:
            for row in range(9, selected['backup_9'].sheet_by_index(0).nrows):
                comparison_list_1.append(selected['backup_9'].sheet_by_index(0).cell_value(row, 2) + selected['backup_9'].sheet_by_index(0).cell_value(row, 3))

        matches_1 = 0
        for concat_str in comparison_list_1:
            if part_and_customer_code == concat_str:
                matches_1 += 1

        if matches_1 >= 1:
            print ('TTC Part No. check 3 --- Pass (Found in Inner Packing BOM)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches_1, 'P/N + Customer Code found in Inner Packing BOM (Submitted/System)')
        else:
            print ('TTC Part No. check 3 --- Fail (P/N + Customer Code not found in Inner Packing BOM)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, matches_1, 'P/N + Customer Code cannot be found in Inner Packing BOM (Submitted/System)')

    # Must be alphanumeric with period or <same as exp>
    # FUTURE: check format
    def customer_parts_imp_hs_code(cell_row, cell_col, new_mod):
        imp_hs_code = ''.join(master_files['xl_sheet_main'].cell_value(cell_row, cell_col).split('.'))

        if imp_hs_code.isalnum():
            print ('Imp HS Code check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), imp_hs_code, 'Imp HS Code is alphanumeric with period')
        else:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col).upper() == '<SAME AS EXP>':
                print ('Imp HS Code check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', '<SAME AS EXP> can be input for Imp HS Code Field')
            else:
                print ('Imp HS Code check --- Fail (Must be alphanumeric with period or <same as exp>)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Must be alphanumeric with period or <SAME AS EXP>')

    # Must be in Region Master and match with Customer Code
    def customer_parts_imp_country(cell_row, cell_col, new_mod):
        imp_country = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        if imp_country == 'I1':
            imp_country = 'IN'
        elif any(imp_country == x for x in ('C1', 'C2', 'C3')):
            imp_country = 'CN'

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in region_master or master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'IN':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-4)[:2].find(imp_country) != -1:
                print ('Imp Country check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-4), 'In Region Master and matches with Customer Code')
            else:
                print ('Imp Country check --- Fail (%s does not match with Customer Code)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-4), 'Does not match with Customer Code')
        else:
            print ('Imp country check --- Fail (%s is not found in region master)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Not found in region master')

    # Ensure input for WEST Imp Countries
    def customer_parts_west_invoice(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6)[:2] in west_import.keys():
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != "":
                print ('WEST Invoice Customer Parts No. check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'Invoice No. input for WEST Imp Country')
            else:
                print ('WEST Invoice Customer Parts No. check --- Fail (Must be input for WEST Imp Countries)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'Invoice No. must be input for WEST Imp Country')
        else:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != "":
                print ('WEST Invoice Customer Parts No. check --- Fail (Should be left blank for WEST Imp Country)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'Should be left blank for non-WEST Imp Country')
            else:
                print ('WEST Invoice Customer Parts No. check --- Pass (Not WEST Imp Country)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'Non-WEST Imp Country field left blank')

    def customer_parts_next_spq(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, 0) == 'MOD' and master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == '':
            return

        spq = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        orderlot = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)

        customer_parts_customer_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-9) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-8)

        # customer_contract_details[Part No. + Customer Code] = Part No. + Supplier Code
        customer_contract_details, supplier_parts = {}, {}
        customer_contract_details_type, supplier_parts_type = '', ''

        try:
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                customer_contract_details[selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 4)] = selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 11)

            for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                customer_contract_details[str(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3)) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 4)] = str(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3)) + additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 11)

            customer_contract_details_type = '1'
        except KeyError:
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                customer_contract_details[selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 4)] = selected['backup_0'].sheet_by_index(0).cell_value(row, 3) + selected['backup_0'].sheet_by_index(0).cell_value(row, 11)

            customer_contract_details_type = '2'

        try:
            for row in range(9, selected['backup_3'].sheet_by_index(0).nrows):
                supplier_parts[selected['backup_3'].sheet_by_index(0).cell_value(row, 2) + selected['backup_3'].sheet_by_index(0).cell_value(row, 3)] = selected['backup_3'].sheet_by_index(0).cell_value(row, 9)

            for row in range(9, additional['TNM_SUPPLIER_PARTS_MASTER'].nrows):
                if additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 10) != '':
                    supplier_parts[str(additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 3)] = additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 10)
                else:
                    supplier_parts[str(additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 3)] = additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 9)

            supplier_parts_type = '1'
        except KeyError:
            for row in range(9, selected['backup_3'].sheet_by_index(0).nrows):
                supplier_parts[selected['backup_3'].sheet_by_index(0).cell_value(row, 2) + selected['backup_3'].sheet_by_index(0).cell_value(row, 3)] = selected['backup_3'].sheet_by_index(0).cell_value(row, 9)

            supplier_parts_type = '2'

        try:
            srbq = supplier_parts[customer_contract_details[customer_parts_customer_code]]

            if srbq % spq == 0 or spq % srbq == 0:
                if orderlot % srbq == 0:
                    print ('SPQ check --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'OL: ' + str(orderlot) + ', SPQ: ' + str(spq), 'SRBQ: ' + str(srbq), 'SPQ:SRBQ = 1:N or N:1, OL:SRBQ = N:1')
                else:
                    print ('SPQ check --- Fail')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'OL: ' + str(orderlot) + ', SPQ: ' + str(spq), 'SRBQ: ' + str(srbq), 'OL:SRBQ != N:1')
            else:
                print ('SPQ check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'OL: ' + str(orderlot) + ', SPQ: ' + str(spq), 'SRBQ: ' + str(srbq), 'SPQ:SRBQ != 1:N or N:1')

        except KeyError:
            print ('SPQ check --- Fail (Corresponding Supplier Part No. cannot be found)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'OL: ' + str(orderlot) + ', SPQ: ' + str(spq), 'NA', 'Corresponding Supplier Part No. cannot be found')

    # Apply date must be 1st of future month if Orderlot is changed.
    def customer_parts_ol_apply_date_new(cell_row, cell_col, new_mod):
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
                print ('Order Lot Apply Date check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', correct_apply_date), 'Apply date is 1st of future month')
            elif apply_date == additional_order_apply_date:
                print ('Order Lot Apply Date check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', additional_order_apply_date), 'Apply date is 1st of Current month')
            else:
                print ('Order Lot Apply Date check --- Fail (must be 1st of next month)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(time.strftime('%d %b %Y', correct_apply_date)) + ' or ' + str(time.strftime('%d %b %Y', additional_order_apply_date)), 'Must be 1st of Next Month or Current Month')

    # IP Specs cannot be zero, cannot be too large (>1)
    def customer_parts_ip_specs(cell_row, cell_col, new_mod):
        box_length = round(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 3)
        box_width = round(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 3)
        box_height = round(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2), 3)

        if box_length == 0 or box_width == 0 or box_height == 0:
            print ('IP Specs check --- Fail (LWH cannot be 0)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'L: ' + str(box_length) + ', W: ' + str(box_width) + ', H: ' + str(box_height), 'NA', "LWH cannot be 0")
        else:
            if box_length >1 or box_width >1 or box_height >1:
                print ('IP Specs check --- Fail (LWH should not be >1)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', 'L: ' + str(box_length) + ', W: ' + str(box_width) + ', H: ' + str(box_height), 'NA', "LWH should not be too large (>1), please check with user")
            else:
                print ('IP Specs check --- Pass (%.3f * %.3f * %.3f)' % (box_length, box_width, box_height))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'L: ' + str(box_length) + ', W: ' + str(box_width) + ', H: ' + str(box_height), 'NA', "LWH not 0, not too large (>1)")

    # IP Gross Weight / SPQ must be more than or equal to Parts Net Weight
    def customer_parts_gross_weight(cell_row, cell_col, new_mod):
        ip_gross_weight = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6) != '':
            spq = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6)
        else:
            spq = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7)

        part_no_net_weight = {}

        for row in range(9, selected['backup_1'].sheet_by_index(0).nrows):
            part_no_net_weight[selected['backup_1'].sheet_by_index(0).cell_value(row, 2)] = selected['backup_1'].sheet_by_index(0).cell_value(row, 7)

        try:
            for row in range(9, additional['TNM_PARTS_MASTER'].nrows):
                part_no_net_weight[additional['TNM_PARTS_MASTER'].cell_value(row, 2)] = additional['TNM_PARTS_MASTER'].cell_value(row, 7)
        except KeyError:
            pass

        try:
            net_weight = part_no_net_weight[master_files['xl_sheet_main'].cell_value(cell_row, cell_col-15)]
        except KeyError:
            print ('Gross Weight check --- Fail (Cannot find P/N)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'GW: ' + str(ip_gross_weight) + ', SPQ: ' + str(spq), 'NA', 'Cannot find P/N in Part Master (Submitted/System)')
            return

        if ip_gross_weight / spq >= net_weight:
            print ('Gross Weight check --- Pass (GW: %.3f, SPQ: %d, NW: %.3f)' % (ip_gross_weight, spq, net_weight))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'GW: ' + str(ip_gross_weight) + ', SPQ: ' + str(spq), str(net_weight), 'IP Gross Weight / SPQ >= Parts Net Weight')
        else:
            print ('Gross Weight check --- Fail (GW/SPQ: %.3f, NW: %.3f)' % (ip_gross_weight / spq, net_weight))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'GW: ' + str(ip_gross_weight) + ', SPQ: ' + str(spq), str(net_weight), 'IP Gross Weight / SPQ must be >= Parts Net Weight')

    # IP Specs apply date for new parts should be <MRS Upload Date>
    def customer_parts_ip_apply_date_new(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col).upper() == '<MRS UPLOAD DATE>':
            print ('IP Specs Apply Date check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Apply date is <MRS UPLOAD DATE>')
        else:
            print ('IP Specs Apply Date check --- Fail (%s must be <MRS Upload Date> for New Parts)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Apply date must be <MRS UPLOAD DATE> for NEW Parts')

    # If paired parts, TTC Part No. of Paired Parts & Paired Parts Flag must be input, P/N must be registered in part master
    def customer_parts_paired_parts(cell_row, cell_col, new_mod):
        paired_parts = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        ttc_part_no_paired = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)
        paired_parts_flag = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)

        if paired_parts == 'Y':
            if ttc_part_no_paired != '' and paired_parts_flag != '':
                part_no_backup = []
                for row in range(9, selected['backup_1'].sheet_by_index(0).nrows):
                    part_no_backup.append(selected['backup_1'].sheet_by_index(0).cell_value(row, 2))

                part_no_backup = list(set(part_no_backup)) # Remove duplicates from Parts Master

                try:
                    for row in range(9, additional['TNM_PARTS_MASTER'].nrows):
                        if additional['TNM_PARTS_MASTER'].cell_value(row, 0) == 'NEW':
                            part_no_backup.append(additional['TNM_PARTS_MASTER'].cell_value(row, 2))
                except KeyError:
                    pass

                if ttc_part_no_paired in part_no_backup:
                    print ('Paired Parts check --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([paired_parts, ttc_part_no_paired, paired_parts_flag]), 'NA', 'TTC P/N must be registered in Parts Master')

                    orderlot = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7)

                    customer_part_no_backup = [] # tuple(P/N, Customer Code, Order Lot)
                    for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
                        if selected['backup_2'].sheet_by_index(0).cell_value(row, 2) == ttc_part_no_paired:
                            customer_part_no_backup.append((selected['backup_2'].sheet_by_index(0).cell_value(row, 2), selected['backup_2'].sheet_by_index(0).cell_value(row, 3), selected['backup_2'].sheet_by_index(0).cell_value(row, 12)))

                    try:
                        for row in range(9, additional['TNM_CUSTOMER_PARTS_MASTER'].nrows):
                            if additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2) == ttc_part_no_paired and additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 0) == 'NEW':
                                customer_part_no_backup.append((additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 12)))
                            elif additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2) == ttc_part_no_paired and additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 0) == 'MOD':
                                for tuple in customer_part_no_backup:
                                    if additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2) == tuple[0] and additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3) == tuple[1]:
                                        customer_part_no_backup.remove(tuple)
                                customer_part_no_backup.append((additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 12)))
                    except KeyError:
                        pass

                    backup_orderlot_list = []
                    for tuple in customer_part_no_backup:
                        backup_orderlot_list.append(tuple[2])

                    validate_count = 0
                    for backup_orderlot in backup_orderlot_list:
                        if orderlot == backup_orderlot:
                            validate_count += 1

                    if validate_count == len(backup_orderlot_list):
                        print ('Paired Parts Orderlot check --- Pass')
                        update_df(new_mod, columns[cell_col+1], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ttc_part_no_paired + ', ' + str(orderlot), ', '.join(str(tuple) for tuple in backup_orderlot_list), 'Order Lot of Paired TTC P/N validated as correct')
                    else:
                        print ('Paired Parts Orderlot check --- Fail')
                        update_df(new_mod, columns[cell_col+1], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ttc_part_no_paired + ', ' + str(orderlot), ', '.join(str(tuple) for tuple in backup_orderlot_list), 'Order Lot of Paired TTC P/N must be the same as P/N Order Lot')

                else:
                    print ('Paired Parts check --- Fail')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ttc_part_no_paired, 'NA', 'TTC P/N must be registered in Parts Master')
            else:
                print ('Paired Parts check --- Fail (Paired Parts field must be input)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', paired_parts, ', '.join([ttc_part_no_paired, paired_parts_flag]), 'Paired parts fields must be input')
        else:
            print ('Paired Parts check --- Pass (Not paired part)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Not a paired part')

    # Back No should be unique per TTC contract, no system restriction
    def customer_parts_back_no(cell_row, cell_col, new_mod):
        # If back no. not input, skip check
        if (any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == x for x in ('', 'N.A.', 'NA', 'N/A', '-', 'Null'))):
            return

        back_no_part_no = []
        for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
            back_no_part_no.append((selected['backup_2'].sheet_by_index(0).cell_value(row, cell_col), selected['backup_2'].sheet_by_index(0).cell_value(row, cell_col-20)))

        comparison_list_1 = []
        for tuple in back_no_part_no:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == tuple[0]:
                comparison_list_1.append(tuple[1])

        if len(comparison_list_1) == 0:
            print ('Back No check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), comparison_list_1, 'Back No. is unique')
        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col-20) in comparison_list_1:
            print ('Back No check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), comparison_list_1, 'Back No. is unique')
        else:
            # Should warn only if same exp country
            supplier_part_exp_country = []
            for part_no in comparison_list_1:
                for row in range(9, selected['backup_3'].sheet_by_index(0).nrows):
                    if selected['backup_3'].sheet_by_index(0).cell_value(row, 2) == part_no:
                        supplier_part_exp_country.append((selected['backup_3'].sheet_by_index(0).cell_value(row, 2), selected['backup_3'].sheet_by_index(0).cell_value(row, 3), selected['backup_3'].sheet_by_index(0).cell_value(row, 8)))

                try:
                    for row in range(9, additional['TNM_SUPPLIER_PARTS_MASTER'].nrows):
                        if additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 0) == 'NEW' and additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 2) == part_no:
                            supplier_part_exp_country.append((additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 2), additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 3), additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 8)))
                        elif additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 0) == 'MOD' and additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 2) == part_no:
                            for tuple in supplier_part_exp_country:
                                if additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 2) == tuple[0] and additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 3) == tuple[1]:
                                    supplier_part_exp_country.remove(tuple)
                            supplier_part_exp_country.append((additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 2), additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 3), additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 8)))
                except KeyError:
                    pass

            print ('Back No check --- Fail (%s is already in use, please confirm with user)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(supplier_part_exp_country)), 'Back No. already in use by another P/N. Please check with user')

    # IPT cannot be less than 1
    def customer_parts_inner_packing_time(cell_row, cell_col, new_mod):
        if int(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) < 1:
            print ('Inner Packing Time check --- Fail (IPT is too short)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Inner Packing Time cannot be <1')
        else:
            print ('Inner Packing Time check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Inner Packing Time not <1')

    # Use colour to test columns to be MOD
    def get_mod_columns(cell_row):
        cols_to_check = []
        for col in range(0, 27): # Hard code MAX COLUMNS

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
    def customer_parts_mod_reference(cell_row):

        # Get concat key
        part_no_customer_code = str(master_files['xl_sheet_main'].cell_value(cell_row, 2)) + master_files['xl_sheet_main'].cell_value(cell_row, 3)

        # Extract all backup concat into list
        comparison_list_1 = []
        for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
            comparison_list_1.append((row, str(selected['backup_2'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_2'].sheet_by_index(0).cell_value(row, 3)))

        # Find backup row
        backup_row = 0
        for concat_str in comparison_list_1:
            if part_no_customer_code == concat_str[1]:
                backup_row = concat_str[0]

        # If cannot find, return False
        if backup_row == 0:
            print ('MOD Reference check --- Fail')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Cannot find Backup Row in system for MOD part')
            return False

        if backup_row >= 0:
            # Extract contents of backup row
            backup_row_contents = []
            for col in range(0, selected['backup_2'].sheet_by_index(0).ncols):
                backup_row_contents.append(selected['backup_2'].sheet_by_index(0).cell_value(backup_row, col))

        submitted_contents = []
        validate_count = 0
        for col in range(0, 27): # Hard Code MAX COLUMN

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

        if validate_count == len(range(2, 27)): # Hard Code MAX COLUMN
            print ('MOD Reference check --- Pass')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', str(submitted_contents), str(backup_row_contents), 'Fields are correctly coloured to indicate \'TO CHANGE\'')

        return True

    # Apply date must be 1st of future month if orderlot is changed, do not change if orderlot is not changed
    # Must be 1st of this month or 1st of next month. -- NEW
    # If 1st of this month, part must not have any (non-cancelled) order in past or current month.
    def customer_parts_ol_apply_date_mod(cell_row, cell_col, new_mod):
        # Get concat key
        part_and_customer_code = master_files['xl_sheet_main'].cell_value(cell_row, 2) + master_files['xl_sheet_main'].cell_value(cell_row, 3)

        concat_orderlot = {}
        for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
            concat_orderlot[selected['backup_2'].sheet_by_index(0).cell_value(row, 2) + selected['backup_2'].sheet_by_index(0).cell_value(row, 3)] = selected['backup_2'].sheet_by_index(0).cell_value(row, 12)

        backup_orderlot = concat_orderlot[part_and_customer_code]

        # If order lot has been changed
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) != backup_orderlot:
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
                    print ('OL Apply Date check 2 --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', correct_apply_date), 'Apply date is 1st of future month')
                elif apply_date == additional_order_apply_date and master_files['xl_sheet_main'].cell_value(cell_row, 0) == 'MOD':
                    print ('OL Apply Date check 2 --- Warning')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', additional_order_apply_date), 'Apply date is 1st of this month, please check if part has non-cancelled order in past or current month')
                elif apply_date == additional_order_apply_date:
                    print ('OL Apply Date check 2 --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', additional_order_apply_date), 'Apply date is 1st of this month')
                else:
                    print ('OL Apply Date check 2 --- Fail (must be 1st of next month)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(time.strftime('%d %b %Y', correct_apply_date)) + ' or ' + str(time.strftime('%d %b %Y', correct_apply_date)), 'Must be 1st of Next Month or Current Month')
        else:
            print ('Order Lot Apply Date check --- Fail (Orderlot is not changed)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Orderlot is not changed')

    # If IP Specs are changed and SPQ is changed, apply date should be 1st of next month to match SPQ reflection.
    # If only IP Specs are changed, apply date can be today or any future date.
    # Please ensure that IP specs Apply Date makes sense with Orderlot Apply Date.
    def customer_parts_ip_apply_date_mod(cell_row, cell_col, new_mod):
        # Get concat key
        part_and_customer_code = master_files['xl_sheet_main'].cell_value(cell_row, 2) + master_files['xl_sheet_main'].cell_value(cell_row, 3)

        # concat_specs_spq[concat_key] = (length, width, height)
        concat_specs_spq = {}
        for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
            concat_specs_spq[selected['backup_2'].sheet_by_index(0).cell_value(row, 2) + selected['backup_2'].sheet_by_index(0).cell_value(row, 3)] = (selected['backup_2'].sheet_by_index(0).cell_value(row, 14), selected['backup_2'].sheet_by_index(0).cell_value(row, 15), selected['backup_2'].sheet_by_index(0).cell_value(row, 16))

        backup_length = concat_specs_spq[part_and_customer_code][0]
        backup_width = concat_specs_spq[part_and_customer_code][1]
        backup_height = concat_specs_spq[part_and_customer_code][2]

        if (master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3) != backup_length or master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2) != backup_width or master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) != backup_height) and master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7) != '':
            # Extract Date
            try:
                apply_date = time.strptime(str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)),"%d %b %Y")
            except ValueError:
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col).upper() == '<MRS UPLOAD DATE>':
                    print ('IP Specs Apply Date check --- Fail (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Date must be input if IP Specs are changed')
                else:
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
                    print ('IP Specs Apply Date check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', correct_apply_date), 'Apply date is 1st of next month')
                # Else if apply date is current month and match with SPQ Apply Date --- Pass
                elif apply_date == additional_order_apply_date and apply_date == time.strptime(str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5)),"%d %b %Y"):
                    print ('IP Specs Apply Date check --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), time.strftime('%d %b %Y', additional_order_apply_date), 'Apply date is 1st of this month and match with SPQ/OL reflection')
                else:
                    print ('IP Specs Apply Date fail --- Fail (must be 1st of next month)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(time.strftime('%d %b %Y', correct_apply_date)) + ' or ' + str(time.strftime('%d %b %Y', additional_order_apply_date)), 'Apply Date should be 1st of current month or next month to match SPQ/OL reflection')

        elif (master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3) != backup_length or master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2) != backup_width or master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) != backup_height) and master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7) == '':
            # Extract Date
            try:
                apply_date = time.strptime(str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)),"%d %b %Y")
            except ValueError:
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col).upper() == '<MRS UPLOAD DATE>':
                    print ('IP Specs Apply Date check --- Fail (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(time.strftime('%d %b %Y', datetime.date.today().timetuple())), 'Date must be today or any future date if IP Specs are changed and SPQ remains unchanged')
                else:
                    print ('IP Specs Apply Date check --- Fail (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Check Date Format')

            if 'apply_date' in locals():
                if apply_date >= datetime.date.today().timetuple():
                    print ('IP Specs Apply Date check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(time.strftime('%d %b %Y', datetime.date.today().timetuple())), 'Apply date is today or future date')
                else:
                    print ('IP Specs Apply Date check --- Fail (Apply Date cannot be before upload date)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(time.strftime('%d %b %Y', datetime.date.today().timetuple())), 'Apply Date cannot be before upload date')
        else:
            print ('IP Specs Apply Date check --- Fail (Do not change if IP specs are not changed)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Do not change if IP specs are not changed')

    # If possible, user should change IPT on screen
    # Next_SPQ must be input if IPT is changed
    def customer_parts_inner_packing_time_mod(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-12) != '':
            print ('Inner Packing Time MOD check --- Pass (User should change IPT on screen)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'User should change IPT on screen')
        else:
            print ('Inner Packing Time MOD check --- Fail (Next SPQ must be input if IPT is changed)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Next SPQ must be input if IPT is changed')

    # Check if company code is in WEST
    def customer_parts_west_part_master(cell_row, cell_col, new_mod):
        # Get concat key
        part_no = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col))

        # Find WEST Fields in system
        backup_imp_west_fields = []
        for row in range(9, selected['backup_1'].sheet_by_index(0).nrows):
            if selected['backup_1'].sheet_by_index(0).cell_value(row, 2) == part_no:
                backup_imp_west_fields.append((selected['backup_1'].sheet_by_index(0).cell_value(row, 15), selected['backup_1'].sheet_by_index(0).cell_value(row, 16), selected['backup_1'].sheet_by_index(0).cell_value(row, 17), selected['backup_1'].sheet_by_index(0).cell_value(row, 18), selected['backup_1'].sheet_by_index(0).cell_value(row, 19)))

        try:
            for row in range(9, additional['TNM_PARTS_MASTER'].nrows):
                if additional['TNM_PARTS_MASTER'].cell_value(row, 2) == part_no:
                    backup_imp_west_fields.append((additional['TNM_PARTS_MASTER'].cell_value(row, 15), additional['TNM_PARTS_MASTER'].cell_value(row, 16), additional['TNM_PARTS_MASTER'].cell_value(row, 17), additional['TNM_PARTS_MASTER'].cell_value(row, 18), additional['TNM_PARTS_MASTER'].cell_value(row, 19)))

        except KeyError:
            pass

        company_code_list = []
        for tuple in backup_imp_west_fields:
            company_code_list.append(tuple[0]) # Check company code only for now

        company_code_list = list(set(company_code_list)) # Remove duplicates

        try:
            import_office_code = west_import[master_files['xl_sheet_main'].cell_value(cell_row, cell_col+5)]['Office Code']
        except KeyError:
            print ('Parts Master WEST check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_no, 'NA', 'Imp Country does not use WEST')
            return

        if import_office_code in company_code_list:
            print ('Parts Master WEST check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_no + ', ' + import_office_code, company_code_list, 'Company Code registered in Parts Master')
        else:
            if import_office_code == 'S556':
                print ('Parts Master WEST check --- Warning')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', part_no + ', ' + import_office_code, company_code_list, 'Company Code not registered in Parts Master (Optional for TW)')
            else:
                print ('Parts Master WEST check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', part_no + ', ' + import_office_code, company_code_list, 'Company Code not registered in Parts Master, please add MOD row in Parts Master to include Company Code')

    def customer_parts_discontinued(cell_row, cell_col):
        part_no_customer_code = PRIMARY_KEY_1 + PRIMARY_KEY_2

        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            if part_no_customer_code == str(selected['backup_0'].sheet_by_index(0).cell_value(row, 3)) + selected['backup_0'].sheet_by_index(0).cell_value(row, 4):
                if selected['backup_0'].sheet_by_index(0).cell_value(row, 8) == 'N':
                    print ('Discontinued check --- Pass')
                    update_df('MOD', columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_no_customer_code, selected['backup_0'].sheet_by_index(0).cell_value(row, 8), 'Part has not been discontinued')
                else:
                    print ('Discontinued check --- Fail')
                    update_df('MOD', columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', part_no_customer_code, selected['backup_0'].sheet_by_index(0).cell_value(row, 8), 'Part has already been discontinued')

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
                customer_parts_duplicate_key(row, 2, 'NEW')
                customer_parts_part_no_1(row, 2, 'NEW')
                customer_parts_part_no_2(row, 2, 'NEW')
                customer_parts_part_no_3(row, 2, 'NEW')
                customer_parts_west_part_master(row, 2, 'NEW')
                customer_parts_imp_hs_code(row, 6, 'NEW')
                customer_parts_imp_country(row, 7, 'NEW')
                customer_parts_west_invoice(row, 9, 'NEW')
                customer_parts_next_spq(row, 11, 'NEW')
                customer_parts_ol_apply_date_new(row, 13, 'NEW')
                customer_parts_ip_specs(row, 14, 'NEW')
                customer_parts_gross_weight(row, 17, 'NEW')
                customer_parts_ip_apply_date_new(row, 18, 'NEW')
                customer_parts_paired_parts(row, 19, 'NEW')
                customer_parts_back_no(row, 22, 'NEW')
                customer_parts_inner_packing_time(row, 23, 'NEW')
            # Conditional for MOD parts
            else:
                paired_parts_check_cycle = 0
                cols_to_check = get_mod_columns(row)
                if len(cols_to_check) != 0:
                    print('User wishes to MOD the following columns:')
                    for col in cols_to_check:
                        print('%s: %s' % (columns[col+2], master_files['xl_sheet_main'].cell_value(row, col+2)))
                    print()

                    if customer_parts_mod_reference(row):
                        check_maximum_length(row, 'MOD')
                        check_compulsory_fields(row, 'MOD')
                        customer_parts_duplicate_key(row, 2, 'MOD')
                        customer_parts_discontinued(row, 2)

                        for col in cols_to_check:
                            # Mod: TTC Parts No., Customer Code
                            if (any(col+2 == x for x in (2, 3, 10))):
                                print ('%s cannot be modded' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'Cannot be modded')
                            # Mod: Imp HS Code
                            if col+2 == 6:
                                customer_parts_imp_hs_code(row, 6, 'MOD')
                            # Mod: Imp Country
                            if col+2 == 7:
                                customer_parts_exp_country(row, 7, 'MOD')
                            # Mod: WEST Invoice
                            if col+2 == 9:
                                customer_parts_west_invoice(row, 9, 'MOD')
                            # Mod: Next_SPQ, Orderlot
                            if (any(col+2 == x for x in (11, 12))):
                                customer_parts_next_spq(row, 11, 'MOD')
                            # Mod: Orderlot Apply Date
                            if col+2 == 13:
                                customer_parts_ol_apply_date_mod(row, 13, 'MOD')
                            # Mod: IP Specs
                            if (any(col+2 == x for x in (14, 15, 16))):
                                customer_parts_ip_specs(row, 14, 'MOD')
                            # Mod: IP Gross Weight
                            if any(col+2 == x for x in (11, 17)):
                                customer_parts_gross_weight(row, 17, 'MOD')
                            # Mod: IP Specs Apply Date
                            if col+2 == 18:
                                customer_parts_ip_apply_date_mod(row, 18, 'MOD')
                            # Mod: Paired Parts
                            if (any(col+2 == x for x in (19, 20, 21))):
                                if paired_parts_check_cycle == 0:
                                    customer_parts_paired_parts(row, 19, 'MOD')
                                    paired_parts_check_cycle += 1
                            # Mod: Back No
                            if col+2 == 22:
                                customer_parts_back_no(row, 22, 'MOD')
                            # Mod: IPT
                            if col+2 == 23:
                                customer_parts_inner_packing_time_mod(row, 23, 'MOD')
                                customer_parts_inner_packing_time(row, 23, 'MOD')
                            # Mod: optional columns
                            if (any(col+2 == x for x in (4, 5, 24, 25, 26))):
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
