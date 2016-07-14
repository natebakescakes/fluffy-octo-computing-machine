import time
from os import listdir

import xlrd
import pandas as pd

from master_data import west_import
from master_data import currency_master
from master_data import payment_terms_master

# Customer Parts - Open required workbooks and check against
def customer_contract(master_files, path):

    # Dictionary of columns
    columns = {
        0: "NEW/MOD",
        1: "Reason for the change",
        2: 'Customer Contract No.',
        3: "West Section",
        4: "West Sales Contract",
        5: "Forward Exchange Position(Sales)",
        6: "Business Type",
        7: "Imp WH No Unpack Flag",
        8: "Sold-to Party",
        9: "Target Month",
        10: "Customer Product Leadtime",
        11: "Currency",
        12: "Payment Terms",
        13: "Imp Warehouse Flag",
        14: "Imp Warehouse Code",
        15: "Packing by Cust PO",
        16: "Discontinue Indicator"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        5: "Module Group Master",
        7: "Customer Contract Master",
        14: "Customer Master"
    }

    def check_newmod_field(cell_row, cell_col):
        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != x for x in ('NEW', 'MOD')):
            print ('NEW/MOD check --- Fail')
            update_df(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Please check NEW/MOD field for whitespace')

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 17))

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
        compulsory_fields = [2, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16]

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', columns[col_index] + ' is a Compulsory Field')

    # Check for duplicate primary keys
    def customer_contract_duplicate_key(cell_row, cell_col, new_mod):
        customer_contract_list = []
        for row in range(10, master_files['xl_sheet_main'].nrows):
            customer_contract_list.append((master_files['xl_sheet_main'].cell_value(row, 0), master_files['xl_sheet_main'].cell_value(row, 2)))

        matches = 0
        for contract_no in customer_contract_list:
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

    # Customer contract should be unique
    # Customer contract should be in this format: CountryCode(2)-CustomerCode(5)-Running(3)
    def customer_contract_no_check(cell_row, cell_col, new_mod):
        comparison_list_1 = []
        for row in range(10, selected['backup_7'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_7'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) not in comparison_list_1:
            print ('Customer Contract No. check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, 'NA', 'Customer Contract No. is unique')
        else:
            backup_row_contents = []
            for row in range(10, selected['backup_7'].sheet_by_index(0).nrows):
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == selected['backup_7'].sheet_by_index(0).cell_value(row, 2):
                    for col in range(0, 17): # Hard Code column range
                        backup_row_contents.append(selected['backup_7'].sheet_by_index(0).cell_value(row, col))

            submitted_row_contents = []
            for col in range(0, 17): # Hard Code column range
                submitted_row_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            discrepancy_reference = []
            for i, cell_value in enumerate(submitted_row_contents):
                if cell_value != backup_row_contents[i] and all(i != x for x in (0, 1, 2)):
                    discrepancy_reference.append(columns[i])

            if len(discrepancy_reference) == 0:
                discrepancy_reference.append('Submitted has no differences from system')

            print ('Customer Contract No. check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, ', '.join(discrepancy_reference), 'Customer Contract No. already exists in system')

        contract_no_split = master_files['xl_sheet_main'].cell_value(cell_row, cell_col).split('-')
        if len(contract_no_split[1]) <= 5 and master_files['xl_sheet_main'].cell_value(cell_row, cell_col).find(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+6)) == 0:
            try:
                int_transform = int(contract_no_split[2])
                print ('Customer Contract No. check 2 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, 'NA', 'Customer Contract No. in correct format and match Customer (Sold-to Party)')
            except (ValueError or IndexError):
                print ('Customer Contract No. check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Please check Customer Contract No. format: (CountryCode(2)-CustomerCode(5)-Running(3))')
        else:
            print ('Customer Contract No. check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Please check Customer Contract No. format: (CountryCode(2)-CustomerCode(5)-Running(3))')

    # If no new parts are registered for a new contract, confirm purpose of registration with user.
    def customer_contract_new_parts(cell_row, cell_col, new_mod):
        try:
            new_parts = []
            for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5):
                    new_parts.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3))

            if len(new_parts) != 0:
                print ('New Parts check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, 'NA', 'P/N in CCD utilising new Customer Contract No.')
            else:
                print ('New Parts check --- Warning')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', PRIMARY_KEY_1, 'NA', 'No parts in CCD using new Customer Contract No.')
        except KeyError:
            print ('New Parts check --- Warning')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', PRIMARY_KEY_1, 'NA', 'No submitted CCD found')

    def customer_contract_west_fields(cell_row, cell_col, new_mod):
        customer_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)
        west_section = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col))

        try:
            west_sales_contract = str(int(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)))
        except ValueError:
            west_sales_contract = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1))

        forward_exchange_position = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2))

        if customer_contract_no[:2] in west_import.keys() and customer_contract_no[:2] != 'TW':
            if any(x != '' for x in (west_section, west_sales_contract, forward_exchange_position)):
                print ('West fields check --- Pass')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_section, west_sales_contract, forward_exchange_position]), customer_contract_no[:2], 'WEST Fields not blank for WEST Imp Country')
            else:
                print ('West fields check --- Fail')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_section, west_sales_contract, forward_exchange_position]), customer_contract_no[:2], 'WEST Fields cannot be blank for WEST Imp Country')
        elif customer_contract_no[:2] == 'TW':
            if any(x != '' for x in (west_section, west_sales_contract, forward_exchange_position)):
                print ('West fields check --- Pass')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_section, west_sales_contract, forward_exchange_position]), customer_contract_no[:2], 'WEST Fields not blank for WEST Imp Country')
            else:
                print ('West fields check --- Fail')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', ', '.join([west_section, west_sales_contract, forward_exchange_position]), customer_contract_no[:2], 'WEST Fields cannot be blank for WEST Imp Country (Optional for TW)')
        else:
            if all(x == '' for x in (west_section, west_sales_contract, forward_exchange_position)):
                print ('West fields check --- Pass')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_section, west_sales_contract, forward_exchange_position]), customer_contract_no[:2], 'WEST Fields blank for non-WEST Imp Country')
            else:
                print ('West fields check --- Fail')
                update_df(new_mod, 'WEST Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_section, west_sales_contract, forward_exchange_position]), customer_contract_no[:2], 'WEST Fields should be blank for non-WEST Imp Country')

        if customer_contract_no[:2] in west_import.keys():
            try:
                if int(west_sales_contract) == 9999999999:
                    print ('WEST sales contract --- WARNING')
                    update_df(new_mod, columns[3], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', west_sales_contract, 'NA', 'Please check if user if data is not to be interfaced to WEST')
            except ValueError:
                pass

    # Case sensitive for business type: Firm, Inventory, Firm-Inventory
    def customer_contract_business_type(cell_row, cell_col, new_mod):
        if any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == x for x in ('Firm', 'Inventory', 'Firm-Inventory')):
            print ('Business Type check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Correct input: Firm, Inventory or Firm-Inventory')
        else:
            print ('Business Type check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Incorrect input, can only be: Firm, Inventory, Firm-Inventory')

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col+7) == 'N':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Firm':
                print ('Business Type check 2 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+8), 'No Imp WH, Business Type = Firm')
            else:
                print ('Business Type check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+8), 'If no Imp WH, business type must be firm')
        else:
            print ('Business Type check 2 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+8), 'Imp WH Flag = Y')

    def customer_contract_no_unpack_1(cell_row, cell_col, new_mod):
        """
        If Imp WH No Unpack Flag is MOD, check if other Customer Contract which
        has same Module Group have different Imp WH No Unpack Flag from this Customer Contract
        If Imp WH No Unpack Flag is MOD, check if other Customer Contract which
        has same Module Group have different Imp WareHouse Code from this Customer Contract
        """
        customer_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5)
        no_unpack_flag = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        imp_warehouse_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+7)

        if new_mod == 'MOD':
            module_group_list = set([])
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                if selected['backup_0'].sheet_by_index(0).cell_value(row, 5) == customer_contract_no and \
                    selected['backup_0'].sheet_by_index(0).cell_value(row, 8) == 'N':
                    module_group_list.add(selected['backup_0'].sheet_by_index(0).cell_value(row, 7))

            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].sheet_by_index(0).nrows):
                    if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5) == customer_contract_no and \
                        additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 8) == 'N':
                        module_group_list.add(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7))
            except KeyError:
                pass

            customer_contract_list = set([])
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                for module_group_code in module_group_list:
                    if selected['backup_0'].sheet_by_index(0).cell_value(row, 7) == module_group_code and \
                        selected['backup_0'].sheet_by_index(0).cell_value(row, 8) == 'N':
                        customer_contract_list.add(selected['backup_0'].sheet_by_index(0).cell_value(row, 5))

            contracts_with_different_flag, contracts_with_different_warehouse_code = set([]), set([])
            for row in range(10, selected['backup_7'].sheet_by_index(0).nrows):
                for contract_no in customer_contract_list:
                    if selected['backup_7'].sheet_by_index(0).cell_value(row, 2) == contract_no:
                        if selected['backup_7'].sheet_by_index(0).cell_value(row, 7) != no_unpack_flag:
                            contracts_with_different_flag.add(selected['backup_7'].sheet_by_index(0).cell_value(row, 2))

                        if selected['backup_7'].sheet_by_index(0).cell_value(row, 14) != imp_warehouse_code:
                            contracts_with_different_warehouse_code.add(selected['backup_7'].sheet_by_index(0).cell_value(row, 7))

                    try:
                        if master_files['xl_sheet_main'].cell_value(row, 2) == contract_no and \
                            master_files['xl_sheet_main'].cell_value(row, 0) == 'MOD':
                            if master_files['xl_sheet_main'].cell_value(row, 7) != no_unpack_flag:
                                contracts_with_different_flag.add(master_files['xl_sheet_main'].cell_value(row, 2))

                            if master_files['xl_sheet_main'].cell_value(row, 14) != imp_warehouse_code:
                                contracts_with_different_warehouse_code.add(master_files['xl_sheet_main'].cell_value(row, 7))
                    except KeyError:
                            pass

            if len(contracts_with_different_flag) > 0:
                print ('Imp WH No Unpack Flag check 1 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', no_unpack_flag, contracts_with_different_flag, 'Other Customer Contracts with same Module Group have different Imp WH No Unpack Flag from this Customer Contract')
            else:
                print ('Imp WH No Unpack Flag check 1 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', no_unpack_flag, contracts_with_different_flag, 'Other Customer Contracts with same Module Group have the same Imp WH No Unpack Flag as this Customer Contract')

            if len(contracts_with_different_warehouse_code) > 0:
                print ('Imp WH No Unpack Flag check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', no_unpack_flag, contracts_with_different_warehouse_code, 'Other Customer Contracts with same Module Group have different Imp Warehouse Code from this Customer Contract')
            else:
                print ('Imp WH No Unpack Flag check 2 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', no_unpack_flag, contracts_with_different_warehouse_code, 'Other Customer Contracts with same Module Group have the same Imp Warehouse Code as this Customer Contract')
        else:
            print ('Imp WH No Unpack Flag check 1/2 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', no_unpack_flag, "NA", 'NEW Customer Contract does not require this check')

    # If MOD, and Imp WH No Unpack Flag = 'Y', other Customer Contract which
    # has same Module Group should have similar Business Type:
    #
    def customer_contract_no_unpack_2(cell_row, cell_col, new_mod):
        """
        If MOD, and Imp wH No Unpack Flag = 'Y', all other Customer Contracts with
        same Module Group Code should have similar Business Type

        e.g. Firm -> Firm
        e.g. Firm-Inventory -> Inventory/Firm-Inventory
        e.g. Inventory -> Inventory/Firm-Inventory
        """
        customer_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5)
        no_unpack_flag = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        business_type = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)

        if new_mod == 'MOD' and no_unpack_flag == 'Y':
            module_group_list = set([])
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                if selected['backup_0'].sheet_by_index(0).cell_value(row, 5) == customer_contract_no and \
                    selected['backup_0'].sheet_by_index(0).cell_value(row, 8) == 'N':
                    module_group_list.add(selected['backup_0'].sheet_by_index(0).cell_value(row, 7))

            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].sheet_by_index(0).nrows):
                    if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5) == customer_contract_no and \
                        additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 8) == 'N':
                        module_group_list.add(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7))
            except KeyError:
                pass

            customer_contract_list = set([])
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                for module_group_code in module_group_list:
                    if selected['backup_0'].sheet_by_index(0).cell_value(row, 7) == module_group_code and \
                        selected['backup_0'].sheet_by_index(0).cell_value(row, 8) == 'N':
                        customer_contract_list.add(selected['backup_0'].sheet_by_index(0).cell_value(row, 5))

            contracts_with_different_business_type = set([])
            for row in range(10, selected['backup_7'].sheet_by_index(0).nrows):
                for contract_no in customer_contract_list:
                    if selected['backup_7'].sheet_by_index(0).cell_value(row, 2) == contract_no:
                        if business_type == 'Firm' and \
                            selected['backup_7'].sheet_by_index(0).cell_value(row, 8) != 'Firm':
                            contracts_with_different_business_type.add(selected['backup_7'].sheet_by_index(0).cell_value(row, 2))

                        if (business_type == 'Inventory' or business_type == 'Firm-Inventory') and \
                            selected['backup_7'].sheet_by_index(0).cell_value(row, 8) == 'Firm':
                            contracts_with_different_business_type.add(selected['backup_7'].sheet_by_index(0).cell_value(row, 2))

                    try:
                        if master_files['xl_sheet_main'].cell_value(row, 2) == contract_no and \
                            master_files['xl_sheet_main'].cell_value(row, 0) == 'MOD':
                            if business_type == 'Firm' and \
                                master_files['xl_sheet_main'].cell_value(row, 8) != 'Firm':
                                contracts_with_different_business_type.add(master_files['xl_sheet_main'].cell_value(row, 2))

                            if (business_type == 'Inventory' or business_type == 'Firm-Inventory') and \
                                master_files['xl_sheet_main'].cell_value(row, 8) == 'Firm':
                                contracts_with_different_business_type.add(master_files['xl_sheet_main'].cell_value(row, 2))
                    except KeyError:
                        pass

            if len(contracts_with_different_business_type) > 0:
                print ('Imp WH No Unpack Flag check 3 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', no_unpack_flag, contracts_with_different_business_type, 'Other Customer Contracts with same Module Group have different Business Type from this Customer Contract')
            else:
                print ('Imp WH No Unpack Flag check 3 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', no_unpack_flag, contracts_with_different_business_type, 'Other Customer Contracts with same Module Group have the same Business Type as this Customer Contract')
        else:
            print ('Imp WH No Unpack Flag check 3 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', no_unpack_flag, "NA", 'NEW Customer Contract or Imp WH No Unpack Flag is "N" does not require this check')

    def customer_contract_no_unpack_3(cell_row, cell_col, new_mod):
        """
        If MOD, and Imp WH No Unpack Flag is 'Y',
        All module groups of the customer contract must be 'Single'
        """
        customer_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5)
        no_unpack_flag = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        if new_mod == 'MOD' and no_unpack_flag == 'Y':
            module_group_list = set([])
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                if selected['backup_0'].sheet_by_index(0).cell_value(row, 5) == customer_contract_no and \
                    selected['backup_0'].sheet_by_index(0).cell_value(row, 8) == 'N':
                    module_group_list.add(selected['backup_0'].sheet_by_index(0).cell_value(row, 7))

            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].sheet_by_index(0).nrows):
                    # Customer Contract cannot be modded in Customer Contract Details
                    if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5) == customer_contract_no and \
                        additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 8) == 'N':
                        module_group_list.add(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7))
            except KeyError:
                pass

            module_group_with_mixed = []
            for module_group_code in module_group_list:
                for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
                    if selected['backup_5'].sheet_by_index(0).cell_value(row, 2) == module_group_code and \
                        selected['backup_5'].sheet_by_index(0).cell_value(row, 5) == 'M':
                        module_group_with_mixed.append(selected['backup_5'].sheet_by_index(0).cell_value(row, 2))

                try:
                    for row in range(10, additional['TNM_MODULE_GROUP'].nrows):
                        if additional['TNM_MODULE_GROUP'].cell_value(row, 2) == module_group_code and \
                            additional['TNM_MODULE_GROUP'].cell_value(row, 0) == 'MOD' and \
                            additional['TNM_MODULE_GROUP'].cell_value(row, 5) == 'S':
                            module_group_with_mixed.remove(module_group_code)
                except KeyError:
                    pass

            if len(module_group_with_mixed) > 0:
                print ('Imp WH No Unpack Flag check 4 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', no_unpack_flag, module_group_with_mixed, 'Not all Module Groups related to this Customer Contract are Single')
            else:
                print ('Imp WH No Unpack Flag check 4 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', no_unpack_flag, module_group_with_mixed, 'All Module Groups related to this Customer Contract are Single')
        else:
            print ('Imp WH No Unpack Flag check 4 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', no_unpack_flag, "NA", 'NEW Customer Contract or Imp WH No Unpack Flag is "N" does not require this check')

    # Customer (Sold to Party) mst be found in Customer Master
    def customer_contract_customer(cell_row, cell_col, new_mod):
        customer_list = []
        for row in range(9, selected['backup_14'].sheet_by_index(0).nrows):
            customer_list.append(selected['backup_14'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in customer_list:
            print ('Customer (Sold-to Party) check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Customer (Sold-to Party) registered in Customer Master')
        else:
            print ('Customer (Sold-to Party) check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Customer (Sold-to Party) cannot be found in Customer Master')

    # Currency must be found in Currency Master
    # All Customer Contracts 1 Customer Code should have the same Currency, except for ID
    def customer_contract_currency(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in currency_master:
            print ('Currency check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Currency registered in Currency Master')
        else:
            print ('Currency check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Currency cannot be found in Currency Master')

        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3)[:2] != x for x in ('ID', 'TW')):
            # Assume Customer (Sold-to Party) is Customer Code
            customer_code_currency = []
            for row in range(10, selected['backup_7'].sheet_by_index(0).nrows):
                customer_code_currency.append((selected['backup_7'].sheet_by_index(0).cell_value(row, 8), selected['backup_7'].sheet_by_index(0).cell_value(row, 11)))

            currency_list_2 = []
            currency_list_2.append(master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            for tuple in customer_code_currency:
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3) == tuple[0]:
                    currency_list_2.append(tuple[1])

            if len(list(set(currency_list_2))) == 1:
                print ('Currency check 2 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(currency_list_2)), 'All customer contracts for 1 customer code have same currency')
            elif len(list(set(currency_list_2))) < 1:
                print ('Currency check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(currency_list_2)), 'Customer code not registered in system')
            else:
                print ('Currency check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(currency_list_2)), 'All customer contracts for 1 customer code should have same currency')
        else:
            print ('Currency check 2 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3)[:2], 'TW and ID do not require Customer Contracts for 1 Customer Code having the same currency')

    # Payment Terms must be found in Payment Terms Master
    def customer_contract_payment_terms(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in payment_terms_master:
            print ('Payment Terms check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Payment Terms registered in Payment Terms Master')
        else:
            print ('Payment Terms check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Payment Terms cannot be found in Payment Terms Master')

    # If no Imp WH, Warehouse Code must be a dummy WH (No Warehouse Flag = 0 in Warehouse Master)
    def customer_contract_warehouse_code(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) == 'N':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[2:] == '99':
                print ('Warehouse Code check 1 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'Warehouse Code is Dummy Warehouse')
            elif any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == x for x in ('BR21', 'BR22', 'BR23', 'BR24')):
                print ('Warehouse Code check 1 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'Warehouse Code is Dummy Warehouse')
            else:
                print ('Warehouse Code check 1 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'if no Imp WH, WH Code must be a dummy WH')

        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) == 'Y':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[2:] == '99':
                print ('Warehouse Code check 1 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'If WH Flag = Y, WH Code should not be Dummy WH')
            elif any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == x for x in ('BR21', 'BR22', 'BR23', 'BR24')):
                print ('Warehouse Code check 1 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'If WH Flag = Y, WH Code should not be Dummy WH')
            else:
                print ('Warehouse Code check 1 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'WH Code is not dummy WH')
        else:
            print ('Warehouse Code check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'Please enter a valid Warehouse Flag')

        if new_mod == 'MOD':
            print ('Warehouse Code check 2 --- Warning')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'If Imp WH is changed, user needs to create Customer Runs for the new WH')

    def customer_contract_discontinue(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y' and new_mod == 'NEW':
            print ('Discontinue Indicator check NEW --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'NEW Contract should not be registered as discontinued')
        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N' and new_mod == 'NEW':
            print ('Discontinue Indicator check NEW --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Discontinue Indicator valid')
        elif new_mod == 'MOD':
            # Check if all non-discontinued parts in CCD for contract no. are being MOD to discontinue in submitted CCD
            # CCD tuple = (part no., contract no., discontinue indicator)
            submitted_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-14)

            customer_contract_details = []
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                customer_contract_details.append((selected['backup_0'].sheet_by_index(0).cell_value(row, 3), selected['backup_0'].sheet_by_index(0).cell_value(row, 5), selected['backup_0'].sheet_by_index(0).cell_value(row, 8)))

            # CCD discontinue = (primary key concat, discontinue indicator)
            contract_no_to_discontinue = []
            for tuple in customer_contract_details:
                if tuple[1] == submitted_contract_no and tuple[2] == 'N':
                    contract_no_to_discontinue.append(tuple[0] + tuple[1])

            submitted_customer_contract_details = []
            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                    if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'MOD':
                        submitted_customer_contract_details.append((additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 8)))
            except KeyError:
                if not contract_no_to_discontinue:
                    print ('Discontinue indicator check MOD --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'All related CCD have already been discontinued')
                else:
                    print ('Discontinue indicator check MOD --- Fail')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), contract_no_to_discontinue, 'The referenced Customer Contract Details must be discontinued')

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
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(len(contract_no_to_discontinue)), str(len(contract_no_mod_discontinue))]), 'Related Customer Contract Details are being discontinued')
            else:
                print ('Discontinue Indicator check MOD --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(len(contract_no_to_discontinue)), str(len(contract_no_mod_discontinue))]), 'Before discontinue, related Customer Contract Details must be discontinued')

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
    def customer_contract_mod_reference(cell_row):

        # Get concat key
        customer_contract_no = str(master_files['xl_sheet_main'].cell_value(cell_row, 2))

        # Extract all backup concat into list
        comparison_list_1 = []
        for row in range(10, selected['backup_7'].sheet_by_index(0).nrows):
            comparison_list_1.append((row, str(selected['backup_7'].sheet_by_index(0).cell_value(row, 2))))

        # Find backup row
        backup_row = 0
        for concat_str in comparison_list_1:
            if customer_contract_no == concat_str[1]:
                backup_row = concat_str[0]

        # If cannot find, return False
        if backup_row == 0:
            print ('MOD Reference check --- Fail')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Cannot find Backup Row in system for MOD part')
            return False

        if backup_row >= 0:
            # Extract contents of backup row
            backup_row_contents = []
            for col in range(0, selected['backup_7'].sheet_by_index(0).ncols):
                backup_row_contents.append(selected['backup_7'].sheet_by_index(0).cell_value(backup_row, col))

        submitted_contents = []
        validate_count = 0
        for col in range(0, 17): # Hard Code MAX COLUMN

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
                        print ('MOD Reference Check --- WARNING (%s BLACK but MOD)' % columns[col])
                        update_df('MOD', columns[col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, col), backup_row_contents[col], 'Field is indicated as \'BLANK\' but different from system')
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

        if validate_count == len(range(2, 17)): # Hard Code MAX COLUMN
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
                if file.find('MRS_ModuleGroup') != -1:
                    selected['backup_5'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_CustomerContract') != -1 and file.find('MRS_CustomerContractDetail') == -1:
                    selected['backup_7'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_Customer') != -1 and file.find('MRS_CustomerContract') == -1 and file.find('MRS_CustomerContractDetail') == -1 and file.find('MRS_CustomerPartsMaster') == -1:
                    selected['backup_14'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
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
                customer_contract_duplicate_key(row, 2, 'NEW')
                customer_contract_no_check(row, 2, 'NEW')
                customer_contract_new_parts(row, 2, 'NEW')
                customer_contract_west_fields(row, 3, 'NEW')
                customer_contract_business_type(row, 6, 'NEW')
                customer_contract_no_unpack_1(row, 7, 'NEW')
                customer_contract_no_unpack_2(row, 7, 'NEW')
                customer_contract_no_unpack_3(row, 7, 'NEW')
                customer_contract_customer(row, 8, 'NEW')
                customer_contract_currency(row, 11, 'NEW')
                customer_contract_payment_terms(row, 12, 'NEW')
                customer_contract_warehouse_code(row, 14, 'NEW')
                customer_contract_discontinue(row, 16, 'NEW')
            # Conditional for MOD parts
            else:
                west_fields_check_cycle, imp_warehouse_check_cycle = 0, 0
                cols_to_check = get_mod_columns(row)
                if len(cols_to_check) != 0:
                    print('User wishes to MOD the following columns:')
                    for col in cols_to_check:
                        print('%s: %s' % (columns[col+2], master_files['xl_sheet_main'].cell_value(row, col+2)))
                    print()

                    if customer_contract_mod_reference(row):
                        check_maximum_length(row, 'MOD')
                        check_compulsory_fields(row, 'MOD')
                        customer_contract_duplicate_key(row, 2, 'MOD')

                        for col in cols_to_check:
                            # Mod: Primary Key, Business Type
                            if (any(col+2 == x for x in (2, 6))):
                                print ('%s cannot be modded' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'Cannot be modded')
                            # Mod: WEST Fields
                            if (any(col+2 == x for x in (3, 4, 5))):
                                if west_fields_check_cycle == 0:
                                    customer_contract_west_fields(row, 3, 'MOD')
                                    west_fields_check_cycle += 1
                            # Mod: Imp WH No Unpack Flag
                            if col+2 == 7:
                                customer_contract_no_unpack_1(row, 7, 'MOD')
                                customer_contract_no_unpack_2(row, 7, 'MOD')
                                customer_contract_no_unpack_3(row, 7, 'MOD')
                            # Mod: Customer (Sold-to Party)
                            if col+2 == 8:
                                customer_contract_customer(row, 8, 'MOD')
                            # Mod: Currency
                            if col+2 == 11:
                                customer_contract_currency(row, 11, 'MOD')
                            # Mod: Payment Terms
                            if col+2 == 12:
                                customer_contract_payment_terms(row, 12, 'MOD')
                            # Mod: Warehouse Flag, Warehouse Code
                            if any(col+2 == x for x in (13, 14)):
                                if imp_warehouse_check_cycle == 0:
                                    customer_contract_warehouse_code(row, 14, 'MOD')
                                    imp_warehouse_check_cycle += 1
                            # Mod: Discontinue indicator
                            if col+2 == 16:
                                customer_contract_discontinue(row, 16, 'MOD')
                            # Mod: optional columns
                            if (any(col+2 == x for x in (9, 10, 15))):
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
