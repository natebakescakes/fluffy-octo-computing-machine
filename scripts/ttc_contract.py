import xlrd, time
import pandas as pd
from os import listdir

from office_master import office_master
from west_data import west_import, west_export
from incoterms_master import incoterms_master

# TTC Contract - Open required workbooks and check against
def ttc_contract(master_files, path):
    # Dictionary of columns
    columns = {
        0: "NEW/MOD",
        1: "Reason for the change",
        2: "TTC Contract No.",
        3: "West Import Section",
        4: "West Import Purchase No.",
        5: "West Export Section",
        6: "West Export Sales No.",
        7: "Import Forward Exchange Position (Purchase)",
        8: "Middle Forward Exchange Position (Sales)",
        9: "Middle Forward Exchange Position (Purchase)",
        10: "Export Forward Exchange Position (Sales)",
        11: "West Middle Sales Section",
        12: "West Middle Purchase Section",
        13: "Middle country Flag",
        14: "Shipping Route",
        15: "Sold-to Party",
        16: "Imp Currency",
        17: "Imp HS Code Output",
        18: "Imp Incoterms",
        19: "Imp Incoterms (Port/City)",
        20: "Imp Consignee",
        21: "Imp Accountee",
        22: "Imp Delivery Address",
        23: "Imp Payment Terms",
        24: "Customer Inventory Flag",
        25: "Exp Office",
        26: "Exp Currency",
        27: "Exp HS Code Output",
        28: "Exp Incoterms",
        29: "Exp Incoterms (Port/City)",
        30: "Exp Consignee",
        31: "Exp Accountee",
        32: "Exp Shipper",
        33: "Exp Payment Terms",
        34: "Exp E-Signature Flag",
        35: "Mid E-Signature Flag",
        36: "Cargo Insurance",
        37: "Discontinue Indicator"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        4: "TTC Contract Master",
        5: "Module Group Master",
        8: "Shipping Calendar Master",
        12: "Currency Master",
        13: "Payment Terms Master",
        14: "Customer Master",
        15: "Supplier Master"
    }

    def check_newmod_field(cell_row, cell_col):
        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != x for x in ('NEW', 'MOD')):
            print ('NEW/MOD check --- Fail')
            update_df(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Please check NEW/MOD field for whitespace')

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 38))

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
        compulsory_fields = [2, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 32, 34, 35, 37]

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', columns[col_index] + ' is a Compulsory Field')

    # Check for duplicate primary key
    def ttc_contract_duplicate_key(cell_row, cell_col, new_mod):
        ttc_contract_list = []
        for row in range(9, master_files['xl_sheet_main'].nrows):
            ttc_contract_list.append((master_files['xl_sheet_main'].cell_value(row, 0), master_files['xl_sheet_main'].cell_value(row, 2)))

        matches = 0
        for contract_no in ttc_contract_list:
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

    # TTC contract must be unique and in this format: Exp ctry(2)Impctry(2)Running no(2)
    def ttc_contract_no_check(cell_row, cell_col, new_mod):
        comparison_list_1 = []
        for row in range(9, selected['backup_4'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_4'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in comparison_list_1:
            backup_row_contents = []
            for row in range(9, selected['backup_4'].sheet_by_index(0).nrows):
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == selected['backup_4'].sheet_by_index(0).cell_value(row, 2):
                    for col in range(0, 38): # Hard Code column range
                        backup_row_contents.append(selected['backup_4'].sheet_by_index(0).cell_value(row, col))

            submitted_row_contents = []
            for col in range(0, 38): # Hard Code column range
                submitted_row_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            discrepancy_reference = []
            for i, cell_value in enumerate(submitted_row_contents):
                if cell_value != backup_row_contents[i] and all(i != x for x in (0, 1, 2)):
                    discrepancy_reference.append(columns[i])

            if len(discrepancy_reference) == 0:
                discrepancy_reference.append('Submitted has no differences from system')

            print ('TTC Contract No. check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, ', '.join(discrepancy_reference), 'TTC Contract No. already found in system')
        else:
            print ('TTC Contract No. check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, 'NA', 'TTC Contract No. is unique')

        contract_no_split = [master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[:2], master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[2:4], master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[4:6]]

        if contract_no_split[0] == EXPORT_COUNTRY and contract_no_split[1] == IMPORT_COUNTRY:
            try:
                int_transform = int(contract_no_split[2])
                print ('TTC Contract No. check 2 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, int_transform, 'TTC Contract No. in correct format and match Imp/Exp Country')
            except ValueError:
                print ('TTC Contract No. check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Incorrect format: ExpCtry(2)ImpCtry(2)Number(2), e.g. THAR02')
        else:
            print ('TTC Contract No. check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Does not match Imp/Exp Country or incorrect format: ExpCtry(2)ImpCtry(2)Number(2), e.g. THAR02')

    # Ensure WEST fields are input for WEST countries
    def ttc_contract_west_fields(cell_row, cell_col, new_mod):
        west_import_section = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        west_import_purchase_no = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1))
        west_export_section = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)
        west_export_sales_no = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3))
        import_forward_exchange = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4)
        middle_forward_exchange_sales = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+5)
        middle_forward_exchange_purchase = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+6)
        export_forward_exchange = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+7)
        west_middle_section_sales = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+8)
        west_middle_section_purchase = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+9)

        # Check WEST Import, not TW
        if IMPORT_COUNTRY in west_import.keys() and IMPORT_COUNTRY != 'TW':
            # west import section not blank, west import purchase no. not blank, import_forward_exchange T, O, S, blank
            if west_import_section != '' and west_import_purchase_no != '' and any(import_forward_exchange == x for x in ('T', 'O', 'S', '')):
                print ('WEST Import check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_import_section, west_import_purchase_no, import_forward_exchange]), IMPORT_COUNTRY, 'WEST Import fields filled in')
            else:
                print ('WEST Import check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_import_section, west_import_purchase_no, import_forward_exchange]), IMPORT_COUNTRY, 'WEST Import fields not filled in properly')
        else:
            if all(x == '' for x in (west_import_section, west_import_purchase_no, import_forward_exchange)):
                print ('WEST Import check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_import_section, west_import_purchase_no, import_forward_exchange]), IMPORT_COUNTRY, 'WEST Import fields blank for non-WEST Imp Country')
            else:
                print ('WEST Import check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_import_section, west_import_purchase_no, import_forward_exchange]), IMPORT_COUNTRY, 'WEST Import fields should be blank for non-WEST Imp Country')

        # Check WEST Export
        if EXPORT_COUNTRY in west_export.keys():
            # west import section not blank, west import purchase no. not blank, import_forward_exchange T, O, S, blank
            if west_export_section != '' and west_export_sales_no != '' and any(export_forward_exchange == x for x in ('T', 'O', 'S', '')):
                print ('WEST Export check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_export_section, west_export_sales_no, export_forward_exchange]), EXPORT_COUNTRY, 'WEST Export fields filled in')
            else:
                print ('WEST Export check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_export_section, west_export_sales_no, export_forward_exchange]), EXPORT_COUNTRY, 'WEST Export fields not filled in properly')
        else:
            if all(x == '' for x in (west_export_section, west_export_sales_no, export_forward_exchange)):
                print ('WEST Export check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_export_section, west_export_sales_no, export_forward_exchange]), EXPORT_COUNTRY, 'WEST Export fields blank for non-WEST Exp Country')
            else:
                print ('WEST Export check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_export_section, west_export_sales_no, export_forward_exchange]), EXPORT_COUNTRY, 'WEST Export fields should be blank for non-WEST Exp Country')

        # Check WEST Middle
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col+10) == 'Y':
            # middle_forward_exchange_sales == T, O, S, blank,
            # middle_forward_exchange_purchase == T, O, S, blank,
            # west_middle_section_sales not blank west_middle_section_purchase not blank
            if west_middle_section_sales != '' and west_middle_section_purchase != '' and any(middle_forward_exchange_sales == x for x in ('T', 'O', 'S', '')) and any(middle_forward_exchange_purchase == x for x in ('T', 'O', 'S', '')):
                print ('WEST Middle check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_middle_section_sales, west_middle_section_purchase, middle_forward_exchange_sales, middle_forward_exchange_purchase]), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+10), 'WEST Middle fields filled in')
            else:
                print ('WEST Middle check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_middle_section_sales, west_middle_section_purchase, middle_forward_exchange_sales, middle_forward_exchange_purchase]), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+10), 'WEST Middle fields not filled in properly')
        else:
            if all(x == '' for x in (west_middle_section_sales, west_middle_section_purchase, middle_forward_exchange_sales, middle_forward_exchange_purchase)):
                print ('WEST Middle check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', ', '.join([west_middle_section_sales, west_middle_section_purchase, middle_forward_exchange_sales, middle_forward_exchange_purchase]), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+10), 'WEST Middle fields blank for no Mid Country contract')
            else:
                print ('WEST Middle check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ', '.join([west_middle_section_sales, west_middle_section_purchase, middle_forward_exchange_sales, middle_forward_exchange_purchase]), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+10), 'WEST Middle fields should be blank for no Mid Country contract')

    # If middle country flag = Y, Exp currency, exp HS code output, exp incoterms, exp incoterms(port/city), exp cosignee, exp accounttee, and exp payment terms must be input
    def ttc_contract_middle_country(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y':
            exp_currency = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+13)
            exp_hs_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+14)
            exp_incoterms = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+15)
            exp_incoterms_port_city = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+16)
            exp_consignee = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+17)
            exp_accountee = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+18)
            exp_payment_terms = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+20)

            exp_fields_list = [exp_currency, exp_hs_code, exp_incoterms, exp_incoterms_port_city, exp_consignee, exp_accountee, exp_payment_terms]

            if all(x != '' for x in exp_fields_list):
                print ('Middle Country Flag check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join(exp_fields_list), 'All optional exp fields are validated as filled in')
            else:
                print ('Middle Country Flag check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join(exp_fields_list), 'All optional exp fields should be filled in if Mid Country Flag = Y')

    # Shipping Route must be found in Shipping Calendar Master
    # Shipping Frequency of Shipping Route and Module Group must have at least 1 overlap
    def ttc_contract_shipping_route(cell_row, cell_col, new_mod):
        shipping_route_list = []
        for row in range(9, selected['backup_8'].sheet_by_index(0).nrows):
            shipping_route_list.append(selected['backup_8'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in shipping_route_list:
            print ('Shipping Route check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Shipping Route found in Shipping Calendar Master')
        else:
            print ('Shipping Route check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Shipping Route not found in Shipping Calendar Master')

        ttc_contract_module_group = []

        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-12) == selected['backup_0'].sheet_by_index(0).cell_value(row, 9):
                ttc_contract_module_group.append((selected['backup_0'].sheet_by_index(0).cell_value(row, 9), selected['backup_0'].sheet_by_index(0).cell_value(row, 7)))

        try:
            for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-12) == additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9):
                    ttc_contract_module_group.append((additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7)))
        except KeyError:
            pass

        if len(ttc_contract_module_group) == 0:
            print ('Shipping Route check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), PRIMARY_KEY_1, 'TTC Contract No. does not match any part in CCD, cannot check Shipping Route Frequency match')
            return

        module_shipping_frequency = {}

        for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
            module_shipping_frequency[selected['backup_5'].sheet_by_index(0).cell_value(row, 2)] = selected['backup_5'].sheet_by_index(0).cell_value(row, 8)

        try:
            for row in range(10, additional['TNM_MODULE_GROUP'].nrows):
                module_shipping_frequency[additional['TNM_MODULE_GROUP'].cell_value(row, 2)] = additional['TNM_MODULE_GROUP'].cell_value(row, 8)
        except KeyError:
            pass

        shipping_calendar = {}
        for row in range(9, selected['backup_8'].sheet_by_index(0).nrows):
            shipping_calendar[selected['backup_8'].sheet_by_index(0).cell_value(row, 2)] = selected['backup_8'].sheet_by_index(0).cell_value(row, 12)

        matches = 0
        module_group_list, shipping_frequency_module_group = [], []
        for tuple in list(set(ttc_contract_module_group)):
            module_group_list.append(tuple[1])
            shipping_frequency_module_group.append(module_shipping_frequency[tuple[1]])
            try:
                if module_shipping_frequency[tuple[1]] == shipping_calendar[master_files['xl_sheet_main'].cell_value(cell_row, cell_col)]:
                    matches += 1
            except KeyError:
                print ('Shipping Route check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(list(set(module_group_list))) + '/' + PRIMARY_KEY_1, 'Module Group Code / TTC Contract No. not found in system or submited masters')
                return

        submitted_display_string = 'Route: ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col) + ' Frequency: ' + shipping_calendar[master_files['xl_sheet_main'].cell_value(cell_row, cell_col)]
        reference_display_string = 'Module No.: ' + str(module_group_list) + ' Frequency: ' + str(shipping_frequency_module_group)

        if matches == len(list(set(ttc_contract_module_group))):
            print ('Shipping Route check 2 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', submitted_display_string, reference_display_string, 'Module Group shipping frequency match with shipping frequency of last ETD of Shipping Route')
        else:
            print ('Shipping Route check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', submitted_display_string, reference_display_string, 'Shipping Routes do not match exactly, please check if they overlap')

    # Sold-to Party must be Imp office code from Office Master, even if there is no Imp business line
    def ttc_contract_sold_to(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in office_master:
            print ('Sold-to Party check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Sold-to Party found in Office Master')
        else:
            print ('Sold-to Party check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Sold-to Party not found in Office Master')

    # Currency must be found in Currency Master
    def ttc_contract_currency(cell_row, cell_col, new_mod):
        currency_list = []
        for row in range(9, selected['backup_12'].sheet_by_index(0).nrows):
            currency_list.append(selected['backup_12'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in currency_list:
            print ('Currency check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Currency registered in Currency Master')
        else:
            print ('Currency check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Currency cannot be found in Currency Master')

    # Incoterm must be found in Incoterms master
    def ttc_contract_incoterms(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in incoterms_master:
            print ('Incoterms check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Incoterms found in Incoterms Master')
        else:
            print ('Incoterms check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Incoterms not found in Incoterms Master')

    # Imp Consignee, Accountee, Delivery address must be found in Customer Master of Imp Country.
    def ttc_contract_imp_consignee(cell_row, cell_col, new_mod):
        imp_consignee = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        imp_accountee = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)
        imp_delivery_address = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)

        customer_list = []
        for row in range(9, selected['backup_14'].sheet_by_index(0).nrows):
            customer_list.append(selected['backup_14'].sheet_by_index(0).cell_value(row, 2))

        if all(x != 'SG:TTSPL' for x in (imp_consignee, imp_accountee, imp_delivery_address)):
            if all(x in office_master for x in (imp_consignee, imp_accountee, imp_delivery_address)):
                print ('Imp Consignee check --- Pass')
                update_df(new_mod, ', '.join([columns[20], columns[21], columns[22]]), cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS',', '.join([imp_consignee, imp_accountee, imp_delivery_address]), 'NA', 'Imp Consignee, Accountee and Delivery Address found in Office Master')
            else:
                if all(x in customer_list for x in (imp_consignee, imp_accountee, imp_delivery_address)) and all(x == IMPORT_COUNTRY for x in (imp_consignee[:2], imp_accountee[:2], imp_delivery_address[:2])):
                    print ('Imp Consignee check --- Pass')
                    update_df(new_mod, ', '.join([columns[20], columns[21], columns[22]]), cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS',', '.join([imp_consignee, imp_accountee, imp_delivery_address]), IMPORT_COUNTRY, 'Imp Consignee, Accountee and Delivery Address found in Customer Master of Imp Country')
                else:
                    print ('Imp Consignee check --- Fail')
                    update_df(new_mod, ', '.join([columns[20], columns[21], columns[22]]), cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL',', '.join([imp_consignee, imp_accountee, imp_delivery_address]), IMPORT_COUNTRY, 'Imp Consignee, Accountee and Delivery Address not found in Customer Master of Imp Country or Office Master')
        else:
            print ('Imp Consignee check --- Fail')
            update_df(new_mod, ', '.join([columns[20], columns[21], columns[22]]), cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL',', '.join([imp_consignee, imp_accountee, imp_delivery_address]), 'NA', 'Imp Consignee, Accountee and Delivery Address should not be SG:TTSPL')

    # Payment Terms must be found in Payment Terms Master
    def ttc_contract_payment_terms(cell_row, cell_col, new_mod):
        payment_terms_list = []
        for row in range(9, selected['backup_13'].sheet_by_index(0).nrows):
            payment_terms_list.append(selected['backup_13'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in payment_terms_list:
            print ('Payment Terms check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Payment Terms registered in Payment Terms Master')
        else:
            print ('Payment Terms check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Payment Terms cannot be found in Payment Terms Master')

    # If Customer Inventory Flag = Y, customer contract cross-dock must be 'N' and Imp WH Flag must be 'Y'
    def ttc_contract_customer_inventory(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y':
            customer_contract_details_list = []
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-22) == selected['backup_0'].sheet_by_index(0).cell_value(row, 9):
                    customer_contract_list.append(selected['backup_0'].sheet_by_index(0).cell_value(row, 5))

            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                    if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-22) == additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9):
                        customer_contract_list.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5))
            except KeyError:
                pass

            customer_contract_list = {}
            for row in range(10, selected['backup_0'].sheet_by_index(0).nrows):
                customer_contract_list[selected['backup_0'].sheet_by_index(0).cell_value(row, 2)] = (selected['backup_0'].sheet_by_index(0).cell_value(row, 7), selected['backup_0'].sheet_by_index(0).cell_value(row, 13))

            cross_dock_list = []
            imp_warehouse_flag = []
            for contract_no in list(set(customer_contract_details_list)):
                cross_dock_list.append(customer_contract_list[contract_list][0])
                imp_warehouse_flag.append(customer_contract_list[1])

            if all(x == 'N' for x in cross_dock_list):
                if all(x == 'Y' for x in imp_warehouse_flag):
                    print ('Customer Inventory Flag check --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(cross_dock_list) + ' ' + str(imp_warehouse_flag), 'All Imp Warehouse Flag = Y, all Cross Dock Flag = N')
                else:
                    print ('Customer Inventory Flag check --- Fail')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(imp_warehouse_flag), 'Not all Imp Warehouse Flag = Y')
            else:
                print ('Customer Inventory Flag check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), str(cross_dock_list), 'Not all Cross Dock Flag = N')

        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N':
            print ('Customer Inventory Flag check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'No check if Customer inventory Flag = N')
        else:
            print ('Customer Inventory Flag check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Fixed: Y or N')

    # Exp Office must be Exp office code from office master, even if there is no exp business lline
    def ttc_contract_exp_office(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in office_master:
            print ('Exp Office check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Exp Office found in Office Master')
        else:
            print ('Exp Office check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Exp Office not found in Office Master')

    # Exp Consignee, Accountee must be found in TTC Office or Customer Master of Import country
    def ttc_contract_exp_consignee(cell_row, cell_col, new_mod):
        exp_consignee = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        exp_accountee = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)

        customer_list = []
        for row in range(9, selected['backup_14'].sheet_by_index(0).nrows):
            customer_list.append(selected['backup_14'].sheet_by_index(0).cell_value(row, 2))

        if exp_consignee in customer_list and exp_consignee[:2] == IMPORT_COUNTRY:
            print ('Exp Consignee check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', exp_consignee, IMPORT_COUNTRY, 'Exp Consignee found in Customer Master of Imp Country')
        else:
            if exp_consignee in office_master:
                print ('Exp Consignee check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', exp_consignee, 'NA', 'Exp Consignee found in Office Master')
            else:
                print ('Exp Consignee check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', exp_consignee, IMPORT_COUNTRY, 'Exp Consignee not found in Office Master nor Customer Master of Imp Country')

        if exp_accountee in office_master:
            print ('Exp Office check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', exp_accountee, 'NA', 'Exp Office found in Office Master')
        else:
            print ('Exp Office check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', exp_accountee, 'NA', 'Exp Office not found in Office Master')

    # Exp Shipper must be found in Supplier Master
    # All parts in 1 Supplier Contract must have the same TTC Contract Shipper
    def ttc_contract_exp_shipper(cell_row, cell_col, new_mod):
        exp_shipper = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        supplier_list = []
        for row in range(9, selected['backup_15'].sheet_by_index(0).nrows):
            supplier_list.append(selected['backup_15'].sheet_by_index(0).cell_value(row, 2))

        if exp_shipper in supplier_list and exp_shipper[:2] == EXPORT_COUNTRY:
            print ('Exp Shipper check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', exp_shipper, EXPORT_COUNTRY, 'Exp Shipper found in Supplier Master of Exp Country')

            ttc_contract_list = []
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                if selected['backup_0'].sheet_by_index(0).cell_value(row, 11) == exp_shipper:
                    ttc_contract_list.append(selected['backup_0'].sheet_by_index(0).cell_value(row, 9))

            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                    if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'NEW' and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 11) == exp_shipper:
                        ttc_contract_list.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9))
            except KeyError:
                pass

            if len(list(set(ttc_contract_list))) == 1 and list(set(ttc_contract_list)) == master_files['xl_sheet_main'].cell_value(cell_row, cell_col-30) == ttc_contract_list[0]:
                print ('Exp Shipper check 2 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', exp_shipper, list(set(ttc_contract_list)), 'All parts in 1 Supplier has same TTC Contract Shipper')
            else:
                print ('Exp Shipper check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', exp_shipper, list(set(ttc_contract_list)), 'All parts in 1 Supplier must have same TTC Contract Shipper')
        else:
            if exp_shipper in office_master:
                print ('Exp Shipper check 1 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', exp_shipper, 'NA', 'Exp Shipper found in Office Master')
            else:
                print ('Exp Shipper check 1 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', exp_shipper, EXPORT_COUNTRY, 'Exp Shipper not found in Office Master nor Supplier Master of Exp Country')

    # Exp E-Signature Flag must be 'Y' if Exp Invoice uses E-Signature
    def ttc_contract_exp_signature(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y':
            print ('Exp E-Signature check --- Warning')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Ensure that Exp Managers\' signatures are registered')
        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N':
            print ('Exp E-Signature check --- Warning')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Should be Y if Exp Invoice uses E-Signature')
        else:
            print ('Exp E-Signature check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Fixed: Y or N')

    # For 2 Country model, Mid E-Signature Flag should be 'N'
    def ttc_contract_mid_signature(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-22) == 'Y':
                print ('Mid E-Signature check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-22), 'For 3 Country Model, E-Signature = Y')
            else:
                print ('Mid E-Signature check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-22), 'For 2 Country Model, E-Signature should be N')
        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-22) == 'N':
                print ('Mid E-Signature check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-22), 'For 2 Country Model, E-Signature = N')
            else:
                print ('Mid E-Signature check --- Warning')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-22), 'Please ensure that there is no E-Signature required')
        else:
            print ('Mid E-Signature check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-22), 'Middle Country Flag is Fixed: Y or N')

    # Cargo Insurance must be registered in Cargo Insurance Master
    # For 2 Country model, should be blank
    def ttc_contract_cargo_insurance(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-23) == 'Y':
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'MSIG':
                print ('Cargo Insurance check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-23), 'Cargo Insurance registered in Cargo Insurance Master, for 3 country model')
            elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == '':
                print ('Cargo Insurance check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Optional field, can be blank')
            else:
                print ('Cargo Insurance check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-23), 'Cargo Insurance not found in Cargo Insurance Master, for 3 country model')
        else:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == '':
                print ('Cargo Insurance check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-23), 'Cargo Insurance is blank for 2 Country Model')
            else:
                print ('Cargo Insurance check --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-23), 'Cargo Insurance should be blank for 2 Country Model')

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
    def ttc_contract_mod_reference(cell_row):

        # Get concat key
        ttc_contract_no = str(master_files['xl_sheet_main'].cell_value(cell_row, 2))

        # Extract all backup concat into list
        comparison_list_1 = []
        for row in range(9, selected['backup_4'].sheet_by_index(0).nrows):
            comparison_list_1.append((row, str(selected['backup_4'].sheet_by_index(0).cell_value(row, 2))))

        # Find backup row
        backup_row = 0
        for concat_str in comparison_list_1:
            if ttc_contract_no == concat_str[1]:
                backup_row = concat_str[0]

        # If cannot find, return False
        if backup_row == 0:
            print ('MOD Reference check --- Fail')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Cannot find Backup Row in system for MOD part')
            return False

        if backup_row >= 0:
            # Extract contents of backup row
            backup_row_contents = []
            for col in range(0, selected['backup_4'].sheet_by_index(0).ncols):
                backup_row_contents.append(selected['backup_4'].sheet_by_index(0).cell_value(backup_row, col))

        submitted_contents = []
        validate_count = 0
        for col in range(0, 38): # Hard Code MAX COLUMN

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

        if validate_count == len(range(2, 38)): # Hard Code MAX COLUMN
            print ('MOD Reference check --- Pass')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', str(submitted_contents), str(backup_row_contents), 'Fields are correctly coloured to indicate \'TO CHANGE\'')

        return True

    def ttc_contract_discontinue(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y' and new_mod == 'NEW':
            print ('Discontinue Indicator check NEW --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'NEW Contract should not be registered as discontinued')
        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N' and new_mod == 'NEW':
            print ('Discontinue Indicator check NEW --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Discontinue Indicator valid')
        elif new_mod == 'MOD':
            # Check if all non-discontinued parts in CCD for contract no. are being MOD to discontinue in submitted CCD
            # CCD tuple = (part no., customer contract no., ttc contract no., discontinue indicator)
            submitted_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-35)

            customer_contract_details = []
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                customer_contract_details.append((selected['backup_0'].sheet_by_index(0).cell_value(row, 3), selected['backup_0'].sheet_by_index(0).cell_value(row, 5), selected['backup_0'].sheet_by_index(0).cell_value(row, 9), selected['backup_0'].sheet_by_index(0).cell_value(row, 8)))

            # CCD discontinue = (primary key concat, discontinue indicator)
            contract_no_to_discontinue = []
            for tuple in customer_contract_details:
                if tuple[2] == submitted_contract_no and tuple[3] == 'N':
                    contract_no_to_discontinue.append(tuple[0] + tuple[1])

            submitted_customer_contract_details = []
            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                    if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'MOD':
                        submitted_customer_contract_details.append((additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 8)))
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
                if tuple[2] == submitted_contract_no and tuple[3] == 'Y':
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
                if file.find('MRS_ShippingCalendar') != -1:
                    selected['backup_8'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('IMS_Currency') != -1:
                    selected['backup_12'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('IMS_PaymentTerms') != -1:
                    selected['backup_13'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_Customer') != -1 and file.find('MRS_CustomerContract') == -1 and file.find('MRS_CustomerContractDetail') == -1 and file.find('MRS_CustomerPartsMaster') == -1:
                    selected['backup_14'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
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
            PRIMARY_KEY_2 = 'NA'

            # Set constant for import and export country
            IMPORT_COUNTRY = master_files['xl_sheet_main'].cell_value(row, 15)[:2]
            EXPORT_COUNTRY = master_files['xl_sheet_main'].cell_value(row, 25)[:2]

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
                ttc_contract_duplicate_key(row, 2, 'NEW')
                ttc_contract_west_fields(row, 3, 'NEW') # Multiple fields
                ttc_contract_middle_country(row, 13, 'NEW')
                ttc_contract_shipping_route(row, 14, 'NEW')
                ttc_contract_sold_to(row, 15, 'NEW')
                ttc_contract_currency(row, 16, 'NEW')
                ttc_contract_incoterms(row, 18, 'NEW')
                ttc_contract_imp_consignee(row, 20, 'NEW') # Multiple fields
                ttc_contract_payment_terms(row, 23, 'NEW')
                ttc_contract_customer_inventory(row, 24, 'NEW')
                ttc_contract_exp_office(row, 25, 'NEW')

                # Only required if Middle Country = 'Y'
                if master_files['xl_sheet_main'].cell_value(row, 13) == 'Y':
                    ttc_contract_currency(row, 26, 'NEW')
                    ttc_contract_incoterms(row, 28, 'NEW')
                    ttc_contract_exp_consignee(row, 30, 'NEW') # Multiple fields
                    ttc_contract_payment_terms(row, 33, 'NEW')

                ttc_contract_exp_shipper(row, 32, 'NEW')
                ttc_contract_exp_signature(row, 34, 'NEW')
                ttc_contract_mid_signature(row, 35, 'NEW')
                ttc_contract_cargo_insurance(row, 36, 'NEW')
                ttc_contract_discontinue(row, 37, 'NEW')

            # Conditional for MOD parts
            else:
                west_fields_check_cycle, imp_consignee_check_cycle, exp_consignee_check_cyle = 0, 0, 0
                cols_to_check = get_mod_columns(row)
                if len(cols_to_check) != 0:
                    print('User wishes to MOD the following columns:')
                    for col in cols_to_check:
                        print('%s: %s' % (columns[col+2], master_files['xl_sheet_main'].cell_value(row, col+2)))
                    print()

                    if ttc_contract_mod_reference(row):
                        check_maximum_length(row, 'MOD')
                        check_compulsory_fields(row, 'MOD')
                        ttc_contract_duplicate_key(row, 2, 'MOD')

                        for col in cols_to_check:
                            # Mod: TTC Contract No., Middle Country Flag
                            if any(col+2 == x for x in (2, 13)):
                                print ('%s cannot be modded' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'Cannot be modded')
                            # Mod: WEST fields
                            if any(col+2 == x for x in (3, 4, 5, 6, 7, 8, 9, 10, 11, 12)):
                                if west_fields_check_cycle == 0:
                                    ttc_contract_west_fields(row, 3, 'MOD')
                                    west_fields_check_cycle += 1
                            # Mod: Shipping Route
                            if col+2 == 14:
                                ttc_contract_shipping_route(row, col+2, 'MOD')
                            # Mod: Sold-to Party
                            if col+2 == 15:
                                ttc_contract_sold_to(row, col+2, 'MOD')
                            # Mod: Imp/Exp Currency
                            if any(col+2 == x for x in (16, 26)):
                                ttc_contract_currency(row, col+2, 'MOD')
                            # Mod: Imp/Exp Incoterms
                            if any(col+2 == x for x in (18, 28)):
                                ttc_contract_incoterms(row, col+2, 'MOD')
                            # Mod: Imp Consignee, Accountee, Delivery to
                            if any(col+2 == x for x in (20, 21, 22)):
                                if imp_consignee_check_cycle == 0:
                                    ttc_contract_imp_consignee(row, 20, 'MOD')
                                    imp_consignee_check_cycle += 1
                            # Mod: Imp/Exp Payment Terms
                            if any(col+2 == x for x in (23, 33)):
                                ttc_contract_payment_terms(row, col+2, 'MOD')
                            # Mod: Customer Inventory Flag
                            if col+2 == 24:
                                ttc_contract_customer_inventory(row, col+2, 'MOD')
                            # Mod: Exp Office
                            if col+2 == 25:
                                ttc_contract_exp_office(row, col+2, 'MOD')
                            # Mod: Exp Consignee, Accountee
                            if any(col+2 == x for x in (30, 31)):
                                if exp_consignee_check_cyle == 0:
                                    ttc_contract_exp_consignee(row, 30, 'MOD')
                                    exp_consignee_check_cyle += 1
                            # Mod: Exp Shipper
                            if col+2 == 32:
                                ttc_contract_exp_shipper(row, col+2, 'MOD')
                            # Mod: Exp E-Signature Flag
                            if col+2 == 34:
                                ttc_contract_exp_signature(row, col+2, 'MOD')
                            # Mod: Mid E-Signature Flag
                            if col+2 == 35:
                                ttc_contract_mid_signature(row, col+2, 'MOD')
                            # Mod: Cargo Insurance
                            if col+2 == 36:
                                ttc_contract_cargo_insurance(row, col+2, 'MOD')
                            # Mod: Cargo Insurance
                            if col+2 == 37:
                                ttc_contract_discontinue(row, col+2, 'MOD')
                            # Mod: optional columns
                            if any(col+2 == x for x in (16, 18, 27, 29)):
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
