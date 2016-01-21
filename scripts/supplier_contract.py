import time
from os import listdir

import xlrd
import pandas as pd

from master_data import west_export

# Supplier Contract - Open required workbooks and check against
def supplier_contract(master_files, path):

    # Dictionary of columns
    columns = {
        0: "NEW/MOD",
        1: "Reason for the change",
        2: 'Supplier Contract No.',
        3: 'Supplier Code',
        4: "West Section",
        5: "West Purchase Contract",
        6: "Forward Exchange Position(Purchase)",
        7: "Currency",
        8: "Payment Terms",
        9: "Exp Warehouse Flag",
        10: "Sub Supplier Code",
        11: "Sub Supplier Plant Code",
        12: "Discontinue Indicator"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        6: "Supplier Contract Master",
        12: "Currency Master",
        13: "Payment Terms Master",
        15: "Supplier Master"
    }

    def check_newmod_field(cell_row, cell_col):
        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != x for x in ('NEW', 'MOD')):
            print ('NEW/MOD check --- Fail')
            update_df(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Please check NEW/MOD field for whitespace')

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 12))

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
        compulsory_fields = [2, 3, 7, 8, 9, 12]

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', columns[col_index] + ' is a Compulsory Field')

    # Check for duplicate primary keys
    def supplier_contract_duplicate_key(cell_row, cell_col, new_mod):
        supplier_contract_list = []
        for row in range(9, master_files['xl_sheet_main'].nrows):
            supplier_contract_list.append((master_files['xl_sheet_main'].cell_value(row, 0), master_files['xl_sheet_main'].cell_value(row, 2)))

        matches = 0
        for contract_no in supplier_contract_list:
            # Check if part no same modifier
            if master_files['xl_sheet_main'].cell_value(cell_row, 0) == contract_no[0]:
                if PRIMARY_KEY_1 == contract_no[1]:
                    matches += 1

        if matches == 1:
            print ('Duplicate Key check --- Pass (Primary key is unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, matches, 'Primary key is unique in submitted master')
        elif matches >1:
            print ('Duplicate Key check --- Fail (Primary key is not unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, matches, 'Primary key is not unique in submitted master')

    # Supplier contract should be unique, match with supplier code
    # Supplier contract should be in this format: CountryCode(2)-SupplierCode(5)-Running(3)
    # TODO: Chekc whether Running No. actually sequential
    def supplier_contract_no_check(cell_row, cell_col, new_mod):
        comparison_list_1 = []
        for row in range(9, selected['backup_6'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_6'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) not in comparison_list_1:
            print ('Supplier Contract No. check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, 'NA', 'Supplier Contract No. is unique')
        else:
            backup_row_contents = []
            for row in range(9, selected['backup_6'].sheet_by_index(0).nrows):
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == selected['backup_6'].sheet_by_index(0).cell_value(row, 2):
                    for col in range(0, 13): # Hard Code column range
                        backup_row_contents.append(selected['backup_6'].sheet_by_index(0).cell_value(row, col))

            submitted_row_contents = []
            for col in range(0, 13): # Hard Code column range
                submitted_row_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            discrepancy_reference = []
            for i, cell_value in enumerate(submitted_row_contents):
                if cell_value != backup_row_contents[i] and all(i != x for x in (0, 1, 2)):
                    discrepancy_reference.append(columns[i])

            if len(discrepancy_reference) == 0:
                discrepancy_reference.append('Submitted has no differences from system')

            print ('Supplier Contract No. check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, ', '.join(discrepancy_reference), 'Supplier Contract No. already exists in system')

        contract_no_split = master_files['xl_sheet_main'].cell_value(cell_row, cell_col).split('-')
        if len(contract_no_split[1]) <= 5 and master_files['xl_sheet_main'].cell_value(cell_row, cell_col).find(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)) == 0:
            try:
                int_transform = int(contract_no_split[2])
                print ('Supplier Contract No. check 2 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 'Supplier Contract No. in correct format and match Supplier Code')
            except (ValueError or IndexError):
                print ('Supplier Contract No. check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Please check Supplier Contract No. format: (CountryCode(2)-Supplier(5)-Running(3))')
        else:
            print ('Supplier Contract No. check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Please check Supplier Contract No. format: (CountryCode(2)-Supplier(5)-Running(3))')

    # If no new parts are registered for a new contract, confirm purpose of registration with user.
    def supplier_contract_new_parts(cell_row, cell_col, new_mod):
        try:
            new_parts = []
            for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 12):
                    new_parts.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3))

            if len(new_parts) != 0:
                print ('New Parts check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, 'NA', 'P/N in CCD utilising new Supplier Contract No.')
            else:
                print ('New Parts check --- Warning')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', PRIMARY_KEY_1, 'NA', 'No parts in CCD using new Supplier Contract No.')
        except KeyError:
            print ('New Parts check --- Warning')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', PRIMARY_KEY_1, 'NA', 'No submitted CCD found')

    # Supplier Code must be found in Supplier Master
    def supplier_contract_supplier_code(cell_row, cell_col, new_mod):
        supplier_list = []
        for row in range(9, selected['backup_15'].sheet_by_index(0).nrows):
            supplier_list.append(selected['backup_15'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in supplier_list:
            print ('Supplier Code check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Supplier Code registered in Supplier Master')
        else:
            print ('Supplier Code check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Supplier Code cannot be found in Supplier Master')

    def supplier_contract_west_fields(cell_row, cell_col, new_mod):
        supplier_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)

        west_section = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        west_purchase_contract = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)
        forward_exchange_position = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)

        if supplier_code[:2] in west_export.keys():
            if all(x != '' for x in (west_section, west_purchase_contract)) and any(forward_exchange_position == x for x in ('T', 'O', 'S', '')):
                print ('West fields check --- Pass')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_section, str(west_purchase_contract), forward_exchange_position]), supplier_code[:2], 'WEST Fields not blank for WEST Exp Country')
            else:
                print ('West fields check --- Fail')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_section, str(west_purchase_contract), forward_exchange_position]), supplier_code[:2], 'WEST Fields cannot be blank for WEST Exp Country')
        else:
            if all(x == '' for x in (west_section, west_purchase_contract, forward_exchange_position)):
                print ('West fields check --- Pass')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_section, str(west_purchase_contract), forward_exchange_position]), supplier_code[:2], 'WEST Fields blank for non-WEST Exp Country')
            else:
                print ('West fields check --- Fail')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_section, str(west_purchase_contract), forward_exchange_position]), supplier_code[:2], 'WEST Fields should be blank for non-WEST Exp Country')

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)[:2] in west_export.keys():
            try:
                if int(west_purchase_contract) == 9999999999:
                    print ('WEST sales contract --- WARNING')
                    update_df(new_mod, columns[3], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', west_purchase_contract, 'NA', 'Please check with user if data is not to be interfaced to WEST')
            except ValueError:
                pass

    # Currency must be found in Currency Master
    def supplier_contract_currency(cell_row, cell_col, new_mod):
        currency_list_1 = []
        for row in range(9, selected['backup_12'].sheet_by_index(0).nrows):
            currency_list_1.append(selected['backup_12'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in currency_list_1:
            print ('Currency check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Currency registered in Currency Master')
        else:
            print ('Currency check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Currency cannot be found in Currency Master')

    # Payment Terms must be found in Payment Terms Master
    def supplier_contract_payment_terms(cell_row, cell_col, new_mod):
        payment_terms_list = []
        for row in range(9, selected['backup_13'].sheet_by_index(0).nrows):
            payment_terms_list.append(selected['backup_13'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in payment_terms_list:
            print ('Payment Terms check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Payment Terms registered in Payment Terms Master')
        else:
            print ('Payment Terms check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Payment Terms cannot be found in Payment Terms Master')

    # Exp WH code must be found in Warehouse Master and match with TTC Office Code
    def supplier_contract_warehouse_code(cell_row, cell_col, new_mod):
        warehouse_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        warehouse_code_list = []
        for row in range(9, selected['backup_6'].sheet_by_index(0).nrows):
            warehouse_code_list.append(selected['backup_6'].sheet_by_index(0).cell_value(row, 9))

        if warehouse_code in warehouse_code_list:
            if warehouse_code[:2] == master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6)[:2]:
                print ('Warehouse Code check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', warehouse_code, master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6)[:2], 'Warehouse Code registered before in Supplier Contract Master, match with Supplier Code')
            else:
                print ('Warehouse Code check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', warehouse_code, master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6)[:2], 'Warehouse Code does not match with Supplier Code')
        else:
            print ('Warehouse Code check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', warehouse_code, 'NA', 'Warehouse Code not registered before in Supplier Contract Master, please check Warehouse Master on user screen')

    # Sub Supplier Code only for TH-TBAS
    def supplier_contract_sub_supplier(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7) == 'TH-TBAS':
            if all(x != '' for x in (master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1))):
                print ('Sub Supplier check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)]), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7), 'Filled for TH-TBAS')
            else:
                print ('Sub Supplier check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)]), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7), 'Only for TH-TBAS supplier')
        else:
            if all(x == '' for x in (master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1))):
                print ('Sub Supplier check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)]), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7), 'Blank fields, not TH-TBAS')
            else:
                print ('Sub Supplier check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)]), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7), 'Should only be filled for TH-TBAS supplier')

    def supplier_contract_discontinue(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y' and new_mod == 'NEW':
            print ('Discontinue Indicator check NEW --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'NEW Contract should not be registered as discontinued')
        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N' and new_mod == 'NEW':
            print ('Discontinue Indicator check NEW --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Discontinue Indicator valid')
        elif new_mod == 'MOD':
            # Check if all non-discontinued parts in CCD for contract no. are being MOD to discontinue in submitted CCD
            # CCD tuple = (part no., contract no., discontinue indicator)
            submitted_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-10)

            customer_contract_details = []
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                customer_contract_details.append((selected['backup_0'].sheet_by_index(0).cell_value(row, 3), selected['backup_0'].sheet_by_index(0).cell_value(row, 12), selected['backup_0'].sheet_by_index(0).cell_value(row, 8)))

            # CCD discontinue = (primary key concat, discontinue indicator)
            contract_no_to_discontinue = []
            for tuple in customer_contract_details:
                if tuple[1] == submitted_contract_no and tuple[2] == 'N':
                    contract_no_to_discontinue.append(tuple[0] + tuple[1])

            submitted_customer_contract_details = []
            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                    if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'MOD':
                        submitted_customer_contract_details.append((additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 12), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 8)))
            except KeyError:
                if not contract_no_to_discontinue:
                    print ('Discontinue indicator check MOD --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'All related CCD have already been discontinued')
                else:
                    print ('Discontinue indicator check MOD --- Fail')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), contract_no_to_discontinue, 'The referenced Customer Contract Details must be discontinued')

            # submitted CCD discontinue = (primary key concat, discontinue indicator)
            contract_no_mod_discontinue = []
            for tuple in submitted_customer_contract_details:
                if tuple[1] == submitted_contract_no and tuple[2] == 'Y':
                    contract_no_to_discontinue.append(tuple[0] + tuple[1])

            # check if both lists have the same primary keys
            if not contract_no_to_discontinue and not contract_no_mod_discontinue:
                print ('Discontinue indicator check MOD --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'All related CCD have already been discontinued')
            elif contract_no_to_discontinue.sort() == contract_no_mod_discontinue.sort():
                print ('Discontinue Indicator check MOD --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([len(contract_no_to_discontinue), len(contract_no_mod_discontinue)]), 'Related Customer Contract Details are being discontinued')
            else:
                print ('Discontinue Indicator check MOD --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([len(contract_no_to_discontinue), len(contract_no_mod_discontinue)]), 'Before discontinue, related Customer Contract Details must be discontinued')

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
    def supplier_contract_mod_reference(cell_row):

        # Get concat key
        supplier_contract_no = str(master_files['xl_sheet_main'].cell_value(cell_row, 2))

        # Extract all backup concat into list
        comparison_list_1 = []
        for row in range(9, selected['backup_6'].sheet_by_index(0).nrows):
            comparison_list_1.append((row, str(selected['backup_6'].sheet_by_index(0).cell_value(row, 2))))

        # Find backup row
        backup_row = 0
        for concat_str in comparison_list_1:
            if supplier_contract_no == concat_str[1]:
                backup_row = concat_str[0]

        # If cannot find, return False
        if backup_row == 0:
            print ('MOD Reference check --- Fail')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Cannot find Backup Row in system for MOD part')
            return False

        if backup_row >= 0:
            # Extract contents of backup row
            backup_row_contents = []
            for col in range(0, selected['backup_6'].sheet_by_index(0).ncols):
                backup_row_contents.append(selected['backup_6'].sheet_by_index(0).cell_value(backup_row, col))

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

        if validate_count == len(range(2, 13)): # Hard Code MAX COLUMN
            print ('MOD Reference check --- Pass')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', str(submitted_contents), str(backup_row_contents), 'Fields are correctly coloured to indicate \'TO CHANGE\'')

        return True

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
                if file.find('MRS_SupplierContract') != -1:
                    selected['backup_6'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('IMS_Currency') != -1:
                    selected['backup_12'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('IMS_PaymentTerms') != -1:
                    selected['backup_13'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_Supplier') != -1 and file.find('MRS_SupplierContract') == -1 and file.find('MRS_SupplierPartsMaster') == -1:
                    selected['backup_15'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
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
                input_to_check = input("Please enter 'Y' to continue, any other key to cancel this master check: ")
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
                supplier_contract_duplicate_key(row, 2, 'NEW')
                supplier_contract_no_check(row, 2, 'NEW')
                supplier_contract_new_parts(row, 2, 'NEW')
                supplier_contract_supplier_code(row, 3, 'NEW')
                supplier_contract_west_fields(row, 4, 'NEW')
                supplier_contract_currency(row, 7, 'NEW')
                supplier_contract_payment_terms(row, 8, 'NEW')
                supplier_contract_warehouse_code(row, 9, 'NEW')
                supplier_contract_sub_supplier(row, 10, 'NEW')
                supplier_contract_discontinue(row, 12, 'NEW')
            # Conditional for MOD parts
            else:
                west_fields_check_cycle, imp_warehouse_check_cycle = 0, 0
                cols_to_check = get_mod_columns(row)
                if len(cols_to_check) != 0:
                    print('User wishes to MOD the following columns:')
                    for col in cols_to_check:
                        print('%s: %s' % (columns[col+2], master_files['xl_sheet_main'].cell_value(row, col+2)))
                    print()

                    if supplier_contract_mod_reference(row):
                        check_maximum_length(row, 'MOD')
                        check_compulsory_fields(row, 'MOD')
                        supplier_contract_duplicate_key(row, 2, 'MOD')

                        for col in cols_to_check:
                            # Mod: Primary Key
                            if col+2 == 2:
                                print ('%s cannot be modded' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'Cannot be modded')
                            # Mod: Supplier code
                            if col+2 == 3:
                                supplier_contract_supplier_code(row, 3, 'MOD')
                            # Mod: WEST Fields
                            if (any(col+2 == x for x in (4, 5, 6))):
                                if west_fields_check_cycle == 0:
                                    supplier_contract_west_fields(row, 4, 'MOD')
                                    west_fields_check_cycle += 1
                            # Mod: Currency
                            if col+2 == 7:
                                supplier_contract_currency(row, 7, 'MOD')
                            # Mod: Payment Terms
                            if col+2 == 8:
                                supplier_contract_payment_terms(row, 8, 'MOD')
                            # Mod: Exp Warehouse Code
                            if col+2 == 9:
                                supplier_contract_warehouse_code(row, 9, 'MOD')
                            # Mod: Sub Supplier Code, Sub Supplier Plant code
                            if any(col+2 == x for x in (10, 11)):
                                if sub_supplier_check_cycle == 0:
                                    supplier_contract_sub_supplier(row, 10, 'MOD')
                                    sub_supplier_check_cycle += 1
                            # Mod: Discontinue indicator
                            if col+2 == 12:
                                supplier_contract_discontinue(row, 12, 'MOD')

                else:
                    print ('There is nothing that is being modded.')
                    update_df('MOD', 'NA', row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Nothing is being modded (No field highlighted RED)')

            print ('-' * 10)
        print('*' * 60)

        # Pandas export to excel
        df = pd.DataFrame(check_dict)
        return df
