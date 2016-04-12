import time
from os import listdir

import xlrd
import pandas as pd

from master_data import office_master

# Customer Contract Details - Open required workbooks and check against
def customer_contract_details(master_files, path):

    # Dictionary of columns
    columns = {
        0: "NEW/MOD",
        1: "Reason for the change",
        2: "Customer Parts Name",
        3: "TTC Parts No.",
        4: "Customer Code",
        5: "Customer Contract",
        6: "TTC-Imp",
        7: "Module Group Code",
        8: "Discontinue Indicator",
        9: "TTC Contract",
        10: "TTC-Exp",
        11: "Supplier Code",
        12: "Supplier Contract",
        13: "Fluctuation Rate(%)",
        14: "Supplier Delivery Pattern",
        15: "Dayweek Specify",
        16: "Week Specify",
        17: "Month Specify",
        18: "Day Specify",
        19: "Exp Remarks",
        20: "End User 1",
        21: "End User 2",
        22: "End User 3",
        23: "End User 4",
        24: "End User 5"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        1: "Parts Master",
        2: "Customer Parts Master",
        3: "Supplier Parts Master",
        4: "TTC Contract",
        5: "Module Group Master",
        6: "Supplier Contract Master",
        7: "Customer Contract Master",
        8: "Shipping Calendar Master",
        9: "Model BOM"
    }

    def check_newmod_field(cell_row, cell_col):
        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != x for x in ('NEW', 'MOD')):
            print ('NEW/MOD check --- Fail')
            update_df(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Please check NEW/MOD field for whitespace')

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 25))

        validate_count = 0
        error_fields = []

        # validate if number
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
        compulsory_fields = list(range(3, 15))

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', columns[col_index] + ' is a Compulsory Field')

    # Check for duplicate primary keys
    def customer_contract_details_duplicate_key(cell_row, cell_col, new_mod):
        concat_list = []
        for row in range(9, master_files['xl_sheet_main'].nrows):
            if master_files['xl_sheet_main'].cell_value(row, 0) != '':
                concat_list.append((master_files['xl_sheet_main'].cell_value(row, 0), str(master_files['xl_sheet_main'].cell_value(row, 3)) + master_files['xl_sheet_main'].cell_value(row, 5)))

        if len(set(concat_list)) == len(concat_list):
            print ('Duplicate Key check --- Pass (Primary key is unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1 + PRIMARY_KEY_2, 'NA', 'Primary key is unique in submitted master')
        else:
            print ('Duplicate Key check --- Fail (Primary key is not unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1 + PRIMARY_KEY_2, 'NA', 'Primary key is not unique in submitted master')

    # Check if TTC P/N registered in Parts Master
    def customer_contract_details_part_no_1(cell_row, cell_col, new_mod):
        comparison_list_1 = []
        for row in range(10, selected['backup_1'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_1'].sheet_by_index(0).cell_value(row,2))

        matches_1 = 0
        for part_no in comparison_list_1:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == part_no:
                matches_1 += 1

        if matches_1 >= 1:
            print('TTC Parts No. check 1 --- Pass (System)')
            update_df(new_mod, columns[cell_col], cell_row, master_files['xl_sheet_main'].cell_value(cell_row, 3), master_files['xl_sheet_main'].cell_value(cell_row, 5), 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_1, 'Registered in Parts Master (System)')
        else:
            comparison_list_2 = []
            try:
                for row in range(9, additional['TNM_PARTS_MASTER'].nrows):
                    comparison_list_2.append(str(additional['TNM_PARTS_MASTER'].cell_value(row,2)))

                # Should move to else block after except
                matches_2 = 0
                for part_no in comparison_list_2:
                    if str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) == part_no:
                        matches_2 += 1

                if matches_2 == 1:
                    print('TTC Parts No. check 1 --- Pass (Submitted)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'Found in Parts Master (Submitted)')
                elif matches_2 > 1:
                    print('TTC Parts No. check 1 --- Fail (Duplicate Parts in Submitted)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'Duplicate Parts in Parts Master (Submitted)')
                else:
                    print('TTC Parts No. check 1 --- Fail (Part not found in System or Submitted master %s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'Part not found in Parts Master (Submitted/System)')

            except KeyError:
                print('TTC Parts No. check 1 --- Fail (Part not found in system)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_1, 'Part not found in Parts Master (System), no submitted parts master')

    # Check if TTC P/N + Customer Contract not already in system
    def customer_contract_details_part_no_2(cell_row, cell_col, new_mod):
        part_and_customer_no = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2))

        comparison_list_1 = []
        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            comparison_list_1.append(str(selected['backup_0'].sheet_by_index(0).cell_value(row,3)) + str(selected['backup_0'].sheet_by_index(0).cell_value(row,5)))

        if part_and_customer_no not in comparison_list_1:
            print('TTC Parts No. check 2 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_and_customer_no, 'NA', 'P/N + Customer Contract not already registered')
        else:
            backup_row_contents = []
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                if part_and_customer_no == str(selected['backup_0'].sheet_by_index(0).cell_value(row, 3)) + selected['backup_0'].sheet_by_index(0).cell_value(row, 5):
                    for col in range(0, 25): # Hard Code column range
                        backup_row_contents.append(selected['backup_0'].sheet_by_index(0).cell_value(row, col))

            submitted_row_contents = []
            for col in range(0, 25): # Hard Code column range
                submitted_row_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            discrepancy_reference = []
            for i, cell_value in enumerate(submitted_row_contents):
                if cell_value != backup_row_contents[i] and all(i != x for x in (0, 1, 3, 5)):
                    discrepancy_reference.append(columns[i])

            if len(discrepancy_reference) == 0:
                discrepancy_reference.append('Submitted has no differences from system')

            print('TTC Parts No. check 2 --- Fail (Part found in system)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', part_and_customer_no, ', '.join(discrepancy_reference), 'P/N + Customer Contract found in system')

    # TODO -----------------------------------------------------------
    def customer_contract_details_shipping_route(cell_row, cell_col):
        ttc_contract_route = {}
        try:
            for row in range(10, additional['TNM_TTC_CONTRACT'].nrows):
                ttc_contract_route[additional['TNM_TTC_CONTRACT'].cell_value(row, 2)] = additional['TNM_TTC_CONTRACT'].cell_value(row, 14)
        except KeyError:
            pass
        finally:
            for row in range(10, selected['backup_4'].sheet_by_index(0).nrows):
                ttc_contract_route[selected['backup_4'].sheet_by_index(0).cell_value(row, 2)] = selected['backup_4'].sheet_by_index(0).cell_value(row, 14)
    # TODO -----------------------------------------------------------

    # TTC-Imp must be registered in TTC Office Master and match with customer
    def customer_contract_details_imp(cell_row, cell_col, new_mod):
        imp_country_office = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[:2]
        imp_country_customer_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1)[:2]
        imp_country_customer_contract = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)[:2]

        if any(imp_country_customer_code == x for x in ('I1', 'I2', 'I3')) and any(imp_country_customer_contract == x for x in ('I1', 'I2', 'I3')):
            imp_country_customer_code = 'IN'
            imp_country_customer_contract = 'IN'

        if any(imp_country_customer_code == x for x in ('C1', 'C2', 'C3')) and any(imp_country_customer_contract == x for x in ('C1', 'C2', 'C3')):
            imp_country_customer_code = 'CN'
            imp_country_customer_contract = 'CN'

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in office_master:
            if imp_country_customer_code == imp_country_office:
                if imp_country_customer_contract == imp_country_office:
                    print ('Imp Office check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) + ', ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'Registered in Office Master and match with Customer')
                else:
                    print ('Imp Office check --- Fail (TTC-Imp does not match with Customer Contract)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) + ', ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'TTC-Imp does not match with Customer Contract')
            else:
                print ('Imp Office check --- Fail (TTC-Imp does not match with Customer)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1) + ', ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2), 'TTC-Imp does not match with Customer Code')
        else:
            print ('Imp Office check --- Fail (TTC-Imp not found in Office Master)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'TTC-Imp not found in Office Master')

    # Check if Customer Code is registered in Customer Contract Master, match with Customer code
    def customer_contract_details_customer_contract_1(cell_row, cell_col, new_mod):
        comparison_list_1 = []
        for row in range(10, selected['backup_7'].sheet_by_index(0).nrows):
            comparison_list_1.append(str(selected['backup_7'].sheet_by_index(0).cell_value(row,2)))

        matches_1 = 0
        for part_no in comparison_list_1:
            if str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) == part_no:
                matches_1 += 1

        # Search for customer code in customer contract no.
        if matches_1 == 1:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col).find(str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1))) != -1:
                print('Customer Contract No. check 1 --- Pass (System %s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'Registered in Customer Contract master, match with Customer Code')
            else:
                print('Customer Contract No. check 1 --- Fail (Customer Code does not match)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'Customer Contract does not match with Customer Code')
        else:
            comparison_list_2 = []
            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT'].nrows):
                    comparison_list_2.append(str(additional['TNM_IMP_CUSTOMER_CONTRACT'].cell_value(row,2)))

                # Should move to else block after except
                matches_2 = 0
                for part_no in comparison_list_2:
                    if str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) == part_no:
                        matches_2 += 1

                if matches_2 == 1:
                    if master_files['xl_sheet_main'].cell_value(cell_row, cell_col).find(str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1))) != -1:
                        print('Customer Contract No. check 1 --- Pass (Submitted %s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'Found in submitted Customer Contract master, match with Customer Code')
                    else:
                        print('Customer Contract No. check 1 --- Fail (Customer Code does not match)')
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'Found in submitted Customer Contract master, does not match with Customer Code')
                elif matches_2 > 1:
                    print('Customer Contract No. check 1 --- Fail (Duplicate Parts in Submitted)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'Duplicate contract in Submitted Customer Contract Master')
                else:
                    print('Customer Contract No. check 1 --- Fail (Part not found in System or Submitted master)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'Part not found in System or Submitted master')

            except KeyError:
                print('Customer Contract No. check 1 --- Fail (Customer Contract not found in system)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Customer Contract not found in system, no submitted Customer Contract master')

    # Check if TTC P/N + Customer Code >= 2 Customer Contract, 2 Exp Countries, 2 Supplier Contracts
    def customer_contract_details_customer_contract_2(cell_row, cell_col, new_mod):
        part_no = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
        customer_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)

        part_and_customer_no = part_no + customer_code

        # Comparison List 1 = Customer Contract
        # Comparison List 2 = Export Countries
        # Comparison List 3 = Supplier Contract
        comparison_list_1 = []
        comparison_list_2 = []
        comparison_list_3 = []

        # Generate list of Customer Contract, Exp Countries, Supplier Contract based on submitted P/N + Customer Code key in system
        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            if part_and_customer_no == str(selected['backup_0'].sheet_by_index(0).cell_value(row,3) + selected['backup_0'].sheet_by_index(0).cell_value(row,4)):
                comparison_list_1.append(str(selected['backup_0'].sheet_by_index(0).cell_value(row,5)))

            if part_and_customer_no == str(selected['backup_0'].sheet_by_index(0).cell_value(row,3) + selected['backup_0'].sheet_by_index(0).cell_value(row,4)):
                comparison_list_2.append(str(selected['backup_0'].sheet_by_index(0).cell_value(row,10)))

            if part_and_customer_no == str(selected['backup_0'].sheet_by_index(0).cell_value(row,3) + selected['backup_0'].sheet_by_index(0).cell_value(row,4)):
                comparison_list_3.append(str(selected['backup_0'].sheet_by_index(0).cell_value(row,12)))

        matches_1, matches_2, matches_3 = 0, 0, 0

        # Match Submitted Customer Contract, Exp Countries, Supplier Contract to generated list
        for customer_contract in comparison_list_1:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2) == customer_contract:
                matches_1 += 1
        for exp_country in comparison_list_2:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col+7) == exp_country:
                matches_2 += 1
        for supplier_contract in comparison_list_3:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col+9) == supplier_contract:
                matches_3 += 1

        if matches_1 == len(comparison_list_1):
            print ('Customer Contract No. check 2 --- Pass (Customer Contract < 2)')
            update_df(new_mod, columns[cell_col+2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_and_customer_no, comparison_list_1, 'For P/N + Customer Contract, Customer Contract < 2')
        else:
            part_customer_contract_discontinue = []
            for row in range(9, master_files['xl_sheet_main'].nrows):
                if master_files['xl_sheet_main'].cell_value(row, 0) == 'MOD' and master_files['xl_sheet_main'].cell_value(row, 3) == part_no and master_files['xl_sheet_main'].cell_value(row, 4) == customer_code and master_files['xl_sheet_main'].cell_value(row, 8) == 'Y':
                    part_customer_contract_discontinue.append((
                        str(master_files['xl_sheet_main'].cell_value(row, 3)),
                        master_files['xl_sheet_main'].cell_value(row, 4),
                        master_files['xl_sheet_main'].cell_value(row, 8)
                    ))

            if len(part_customer_contract_discontinue) == 1:
                print ('Customer Contract No. check 2 --- Pass (Discontinue MOD in submitted)')
                update_df(new_mod, columns[cell_col+2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_and_customer_no, comparison_list_1, 'For P/N + Customer Code, Customer Contract > 1, Discontinue MOD found in submitted Customer Contract Details')
            else:
                print ('Customer Contract No. check 2 --- Fail (>1 Customer Contract)')
                update_df(new_mod, columns[cell_col+2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', part_no + customer_code, part_customer_contract_discontinue, 'For P/N + Customer Code, Customer Contract > 1, Discontinue MOD not found in submitted Customer Contract Details')

        if matches_2 == len(comparison_list_2):
            print ('Customer Contract No. check 3 --- Pass (Exp Country < 2)')
            update_df(new_mod, columns[cell_col+2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_and_customer_no, comparison_list_2, 'For P/N + Customer Code, Exp Country < 2')
        else:
            print ('Customer Contract No. check 3 --- Fail (>1 Exp Country)')
            update_df(new_mod, columns[cell_col+2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', part_and_customer_no, comparison_list_2, 'For P/N + Customer Code, Exp Country > 1')

        if matches_3 == len(comparison_list_3):
            print ('Customer Contract No. check 4 --- Pass (Supplier Contract < 2)')
            update_df(new_mod, columns[cell_col+2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_and_customer_no, comparison_list_3, 'For P/N + Customer Code, Supplier Contract < 2')
        else:
            print ('Customer Contract No. check 4 --- Fail (>1 Supplier Contract)')
            update_df(new_mod, columns[cell_col+2], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_and_customer_no, comparison_list_3, 'For P/N + Customer Code, Supplier Contract > 1')

    def customer_contract_details_module_group(cell_row, cell_col, new_mod):
        # Module group must be found in Module Group Master
        comparison_list_1 = []
        for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
            comparison_list_1.append(str(selected['backup_5'].sheet_by_index(0).cell_value(row,2)))

        matches_1 = 0
        for module_group_code in comparison_list_1:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == module_group_code:
                matches_1 += 1

        if matches_1 == 1:
            print('Module Group Code check 1 --- Pass (System %s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_1, 'Module group registered in system')
        else:
            comparison_list_2 = []
            try:
                for row in range(10, additional['TNM_MODULE_GROUP'].nrows):
                    comparison_list_2.append(str(additional['TNM_MODULE_GROUP'].cell_value(row,2)))

                # Should move to else block after except
                matches_2 = 0
                for module_group_code in comparison_list_2:
                    if (master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) == module_group_code:
                        matches_2 += 1

                if matches_2 == 1:
                    print('Module Group Code check 1 --- Pass (Submitted %s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'Module group found in Submitted Module Group Master')
                elif matches_2 > 1:
                    print('Module Group Code check 1 --- Fail (Duplicate Module Group in Submitted)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'Duplicate module group found in Submitted Module Group Master')
                else:
                    print('Module Group Code check 1 --- Fail (Module Group not found in System or Submitted master)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'Not registered in system or found in submitted Module Group Master')

            except KeyError:
                print('TTC Parts No. check 1 --- Fail (Part not found in system)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Not registered in system and no submitted Module Group Master')

        # If Customer Contract cross dock flag = Y, all parts in the module group must be from the same customer
        cross_dock = {}
        try:
            for row in range(10, additional['TNM_IMP_CUSTOMER_CONTRACT'].nrows):
                cross_dock[additional['TNM_IMP_CUSTOMER_CONTRACT'].cell_value(row, 2)] = additional['TNM_IMP_CUSTOMER_CONTRACT'].cell_value(row, 7)
        except KeyError:
            pass
        finally:
            for row in range(10, selected['backup_7'].sheet_by_index(0).nrows):
                cross_dock[selected['backup_7'].sheet_by_index(0).cell_value(row, 2)] = selected['backup_7'].sheet_by_index(0).cell_value(row, 7)

        try:
            if cross_dock[master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)] == 'Y':
                customer_code = []
                for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                    if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == selected['backup_0'].sheet_by_index(0).cell_value(row, 7):
                        customer_code.append(selected['backup_0'].sheet_by_index(0).cell_value(row, 4))

                matches_3 = 0
                for code in customer_code:
                    if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3) == code:
                        matches_3 += 1

                if matches_3 == len(customer_code):
                    print ('Module Group Code check 2 --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(customer_code)), 'Crossdock = Y, Parts in Module Group all same Customer Code')
                else:
                    print ('Module Group Code check 2 --- Fail (Not same customer in crossdock module group)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(customer_code)), 'Crossdock = Y, Parts in Module Group not the same Customer Code')
            else:
                print ('Module Group Code check 2 --- Pass (%s Crossdock Flag == N)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Crossdock = N, no check required')
        except KeyError:
            print ('Module Group Code check 2 --- Fail (Customer Contract not found)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Customer Contract not found in Customer Contract Master or system')

        # If Module Group is single, Module Group Code customer must match with Customer Code
        module_group_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        customer_code_submitted = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3)

        ms_flag = ()
        for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
            if selected['backup_5'].sheet_by_index(0).cell_value(row, 2) == module_group_code:
                ms_flag = (
                    selected['backup_5'].sheet_by_index(0).cell_value(row, 5),
                    selected['backup_5'].sheet_by_index(0).cell_value(row, 6)
                )

        try:
            for row in range(10, additional['TNM_MODULE_GROUP'].nrows):
                if additional['TNM_MODULE_GROUP'].cell_value(row, 2) == module_group_code:
                    ms_flag = (
                        additional['TNM_MODULE_GROUP'].cell_value(row, 5),
                        additional['TNM_MODULE_GROUP'].cell_value(row, 6)
                    )
        except KeyError:
            pass

        if ms_flag == ():
            print ('Module Group Code check 3 --- Fail (Module Group Code not found)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', module_group_code, 'NA', 'Module Group Code not found in system or submitted')
            return

        if ms_flag[0] == 'S':
            if customer_code_submitted == ms_flag[1]:
                print ('Module Group Code check 3 --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', module_group_code, ms_flag, 'MS Flag = S, Module Group Code customer match with Customer Code')
            else:
                print ('Module Group Code check 3 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', module_group_code, ms_flag, 'MS Flag = S, Module Group Code customer discrepancy')
        else:
            print ('Module Group Code check 3 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', module_group_code, ms_flag, 'MS Flag = M, no check required')

    def customer_contract_details_discontinue_new(cell_row, cell_col, new_mod):
        # Check if discontinue indicator is 'N'
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N':
            print ('Discontinue Indicator check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Discontinue indicator is N')
        else:
            print ('Discontinue Indicator check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Discontinue indicator not N')

    def customer_contract_details_ttc_contract_1(cell_row, cell_col, new_mod):
        # TTC Contract No. must be found in TTC contract master *and match with Imp & Exp Country*
        comparison_list_1 = []
        for row in range(9, selected['backup_4'].sheet_by_index(0).nrows):
            comparison_list_1.append(selected['backup_4'].sheet_by_index(0).cell_value(row, 2))

        matches_1 = 0
        for contract_no in comparison_list_1:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == contract_no:
                matches_1 += 1

        if matches_1 == 1:
            print ('TTC Contract No. check 1 --- Pass (System %s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_1, 'TTC Contract No. registered in system')
        else:
            comparison_list_2 = []
            try:
                for row in range(9, additional['TNM_TTC_CONTRACT'].nrows):
                    comparison_list_2.append(str(additional['TNM_TTC_CONTRACT'].cell_value(row,2)))

                # Should move to finally block after except
                matches_2 = 0
                for part_no in comparison_list_2:
                    if str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)) == part_no:
                        matches_2 += 1

                if matches_2 == 1:
                    print('TTC Contract No. check 1 --- Pass (%s Submitted)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'TTC Contract No. found in submitted TTC Contract Master')
                elif matches_2 > 1:
                    print('TTC Contract No. check 1 --- Fail (Duplicate contract in Submitted)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'Duplicate TTC Contract No. found in submitted TTC Contract Master')
                else:
                    print('TTC Contract No. check 1 --- Fail (TTC Contract No. not in System or Submitted master)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), matches_2, 'TTC Contract No. not registered in system or submitted TTC Contract Master')

            except KeyError:
                print('TTC Contract No. check 1 --- Fail (TTC Contract No. not in system)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'TTC Contract No. not registered in system, no TTC Contract Master submitted')

    # Module Group shipping frequency must match with Shipping Calendar frequency
    # Note: This solution iterates to the last date range
    def customer_contract_details_ttc_contract_2(cell_row, cell_col, new_mod):
        module_group_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)
        ttc_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        module_shipping_frequency = ''
        for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
            if selected['backup_5'].sheet_by_index(0).cell_value(row, 2) == module_group_code:
                module_shipping_frequency = selected['backup_5'].sheet_by_index(0).cell_value(row, 8)

        try:
            for row in range(9, additional['TNM_MODULE_GROUP'].nrows):
                # If in submitted, regardless of NEW or MOD, should replace backup
                if additional['TNM_MODULE_GROUP'].cell_value(row, 2) == module_group_code:
                    module_shipping_frequency = additional['TNM_MODULE_GROUP'].cell_value(row, 8)
        except KeyError:
            pass

        if module_shipping_frequency == '':
            print ('TTC Contract No. check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ttc_contract_no, module_group_code, 'Module Group Code cannot be found in Module Group Master')
            return

        shipping_route = ''
        for row in range(9, selected['backup_4'].sheet_by_index(0).nrows):
            if selected['backup_4'].sheet_by_index(0).cell_value(row, 2) == ttc_contract_no:
                shipping_route = selected['backup_4'].sheet_by_index(0).cell_value(row, 14)

        try:
            for row in range(9, additional['TNM_TTC_CONTRACT'].nrows):
                # If in submitted, regardless of NEW or MOD, should replace backup
                if additional['TNM_TTC_CONTRACT'].cell_value(row, 2) == ttc_contract_no:
                    shipping_route = additional['TNM_TTC_CONTRACT'].cell_value(row, 14)
        except KeyError:
            pass

        if shipping_route == '':
            print ('Shipping Frequency check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ttc_contract_no, ttc_contract_no, 'TTC Contract cannot be found in TTC Contract Master')
            return

        shipping_calendar_frequency = ''
        for row in range(9, selected['backup_8'].sheet_by_index(0).nrows):
            # Loop will end with last ETD
            if selected['backup_8'].sheet_by_index(0).cell_value(row, 2) == shipping_route:
                shipping_calendar_frequency = selected['backup_8'].sheet_by_index(0).cell_value(row, 12)

        if shipping_calendar_frequency == '':
            print ('Shipping Frequency check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', ttc_contract_no, shipping_route, 'Shipping Route cannot be found in Shipping Calendar Master')
            return

        # Split frequency into elements, convert to set
        try:
            shipping_calendar_frequency_split = set(shipping_calendar_frequency.split(','))
        except AttributeError:
            shipping_calendar_frequency_split = {int(shipping_calendar_frequency)}

        try:
            module_shipping_frequency_split = set(module_shipping_frequency.split(','))
        except AttributeError:
            module_shipping_frequency_split = {str(int(module_shipping_frequency))}

        if len(set.intersection(shipping_calendar_frequency_split, module_shipping_frequency_split)) != 0:
            print ('Shipping Frequency check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', module_shipping_frequency, shipping_calendar_frequency, 'Module Group shipping frequency intersects with shipping frequency of last ETD of Shipping Route')
        else:
            print ('Shipping Frequency check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', module_shipping_frequency, shipping_calendar_frequency, 'Module Group shipping frequency does not intersect with shipping frequency of last ETD of Shipping Route, please check if date overlaps with other ETDs')

    # All parts in 1 Module Group must have the same TTC Contract No
    def customer_contract_details_ttc_contract_3(cell_row, cell_col, new_mod):
        module_group_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)
        ttc_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        customer_contract_details = []
        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            if selected['backup_0'].sheet_by_index(0).cell_value(row, 7) == module_group_code and selected['backup_0'].sheet_by_index(0).cell_value(row, 8) == 'N':
                customer_contract_details.append((selected['backup_0'].sheet_by_index(0).cell_value(row, 3), selected['backup_0'].sheet_by_index(0).cell_value(row, 5), selected['backup_0'].sheet_by_index(0).cell_value(row, 9)))

        try:
            for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7) == module_group_code and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'NEW':
                    customer_contract_details.append((additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9)))
                elif additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7) == module_group_code and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'MOD':
                    for entry in customer_contract_details:
                        if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3) == entry[0] and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5) == entry[1]:
                            customer_contract_details.remove(entry)
                    customer_contract_details.append((additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 3), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 5), additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9)))
        except KeyError:
            pass



        ttc_contract_list = []
        for entry in customer_contract_details:
            ttc_contract_list.append(entry[2])

        if len(list(set(ttc_contract_list))) == 1:
            print ('TTC Contract No. check 3 --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(ttc_contract_list)), 'All parts in 1 Module Group Code have same TTC Contract No.')
        else:
            print ('TTC Contract No. check 3 --- Fail (Different TTC Contract No. in same Module Group)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(ttc_contract_list)), 'All parts in 1 Module Group Code do not have same TTC Contract No.')

    # For PK, 1 customer has 1 TTC contract
    def customer_contract_details_ttc_contract_pk(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col).find('PK') != -1:
            ttc_contract_customer = []
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                ttc_contract_customer.append((selected['backup_0'].sheet_by_index(0).cell_value(row, 10), selected['backup_0'].sheet_by_index(0).cell_value(row, 4)))

            customer_list = []
            for element in ttc_contract_customer:
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == element[0]:
                    customer_list.append(element[1])

            matches_1 = 0
            for customer_code in customer_list:
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-7) == customer_code:
                    matches_1 += 1

            if matches_1 == len(customer_list):
                print ('TTC Contract No. check 4 --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(customer_list)), 'For PK, 1 customer has 1 TTC Contract No.')
            else:
                print ('TTC Contract no. check 4 --- Fail (1 customer per PK TTC Contract)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), list(set(customer_list)), '1 customer multiple TTC Contract No.')
        else:
            print ('TTC Contract no. check 4 --- Pass (%s Not PK Contract)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Not PK Contract')

    # TTC-Exp must be registered in TTC Office Master and match with customer
    def customer_contract_details_exp(cell_row, cell_col, new_mod):
        exp_country_office = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[:2]
        exp_country_supplier_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)[:2]
        exp_country_supplier_contract = master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)[:2]

        if any(exp_country_supplier_code == x for x in ('C1', 'C2', 'C3')) and any(exp_country_supplier_contract == x for x in ('C1', 'C2', 'C3')):
            exp_country_supplier_code = 'CN'
            exp_country_supplier_contract = 'CN'

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in office_master:
            if exp_country_office == exp_country_supplier_code:
                if exp_country_office == exp_country_supplier_contract:
                    print ('Exp Office check --- Pass (%s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1) + ', ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2), 'Registered in Office Master and match with Supplier')
                else:
                    print ('Exp Office check --- Fail (TTC-Exp does not match with Supplier Contract)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1) + ', ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2), 'TTC-Exp does not match with Supplier Contract')
            else:
                print ('Exp Office check --- Fail (TTC-Exp does not match with Supplier)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1) + ', ' + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2), 'TTC-Exp does not match with Supplier Code')
        else:
            print ('Exp Office check --- Fail (TTC-Exp not found in Office Master)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'TTC-Exp not found in Office Master')

    def customer_contract_details_supplier_code(cell_row, cell_col, new_mod):
        supplier_code = str(master_files['xl_sheet_main'].cell_value(cell_row, 3))
        part_no_supplier_code = supplier_code + str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col))

        supplier_parts = []
        for row in range(9, selected['backup_3'].sheet_by_index(0).nrows):
            if str(selected['backup_3'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_3'].sheet_by_index(0).cell_value(row, 3) == part_no_supplier_code:
                supplier_parts.append(str(selected['backup_3'].sheet_by_index(0).cell_value(row, 2)) + selected['backup_3'].sheet_by_index(0).cell_value(row, 3))

        try:
            for row in range(9, additional['TNM_SUPPLIER_PARTS_MASTER'].nrows):
                if str(additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 3) == part_no_supplier_code and additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 0) == 'NEW':
                    supplier_parts.append(str(additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 2)) + additional['TNM_SUPPLIER_PARTS_MASTER'].cell_value(row, 3))
        except KeyError:
            pass

        if part_no_supplier_code in supplier_parts:
            print ('Supplier Code check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', supplier_code, 'NA', 'Supplier Part not registered in Supplier Parts Master')
        else:
            print ('Supplier Code check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', supplier_code, supplier_parts, 'Supplier Part not registered in Supplier Parts Master')

    # Check if Customer Code is registered in Supplier Contract Master, match with Supplier code
    def customer_contract_details_supplier_contract(cell_row, cell_col, new_mod):
        comparison_list = []
        for row in range(9, selected['backup_6'].sheet_by_index(0).nrows):
            if selected['backup_6'].sheet_by_index(0).cell_value(row, 12) == 'N': # If Discontinue Indicator = N
                comparison_list.append(selected['backup_6'].sheet_by_index(0).cell_value(row, 2))

        try:
            for row in range(9, additional['TNM_EXP_SUPPLIER_CONTRACT'].nrows):
                if additional['TNM_EXP_SUPPLIER_CONTRACT'].cell_value(row, 0) == 'NEW':
                    comparison_list.append(additional['TNM_EXP_SUPPLIER_CONTRACT'].cell_value(row, 2))
                else:
                    # Remove MOD to discontinue parts
                    if additional['TNM_EXP_SUPPLIER_CONTRACT'].cell_value(row, 12) == 'Y':
                        for contract_no in comparison_list:
                            if additional['TNM_EXP_SUPPLIER_CONTRACT'].cell_value(row, 2) == contract_no:
                                comparison_list.remove(contract_no)
        except KeyError:
            pass

        # Search for supplier code in supplier contract no.
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in comparison_list:
            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col).find(str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1))) != -1:
                print('Supplier Contract No. check 1 --- Pass (System %s)' % master_files['xl_sheet_main'].cell_value(cell_row, cell_col))
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'Registered in Supplier Contract Master, match with Supplier Code')
            else:
                print('Supplier Contract No. check 1 --- Fail (Supplier Code does not match)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-1), 'Supplier Contract does not match with Supplier Code')
        else:
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Contract No. not found in Supplier Contract Master (Submitted/System)')

    # Match with input in Day Week Specify, Week Specify, Month Specify, Day Specify
    def customer_contract_details_supplier_delivery_pattern(cell_row, cell_col, new_mod):
        input_to_check = str(int(master_files['xl_sheet_main'].cell_value(cell_row, cell_col)))
        if (any(input_to_check == x for x in ('1', '2', '3', '4', '5'))):
            if input_to_check == '1':
                if str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4)) == '':
                    print ('Supplier Delivery Pattern check --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'Supplier Delivery Pattern filled correctly')
                else:
                    print ('Supplier Delivery pattern check --- Fail (Other fields not blank)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'DayWeek Specify, Week Specify, Month Specify or Day Specify not blank')
            elif input_to_check == '2':
                day_array = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)).split(',')
                matches_1 = 0
                for day in day_array:
                    if (any(day == x for x in ('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'))):
                        matches_1 += 1

                if matches_1 == len(day_array):
                    if str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4)) == '':
                        print ('Supplier Delivery Pattern Check --- Pass')
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'Supplier Delivery Pattern filled correctly')
                    else:
                        print ('Supplier Delivery Pattern Check --- Fail (Other fields not blank)')
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'Week Specify, Month Specify or Day Specify not blank')
                else:
                    print ('Supplier Delivery Pattern Check --- Fail (Check DayWeek Specify)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1), 'NA', 'Mon,Tue,Wed,Thu,Fri,Sat separated by comma')

            elif input_to_check == '3':
                num_array = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)).split(',')
                matches_2 = 0
                for num in num_array:
                    if (any(num == x for x in ('1', '2', '3', '4', '5'))):
                        matches_2 += 1

                if matches_2 == len(num_array):
                    if str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4)) == '':
                        print ('Supplier Delivery Pattern check --- Pass')
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'Supplier Delivery Pattern filled correctly')
                    else:
                        print ('Supplier Delivery Pattern check --- Fail (Other fields not blank)')
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'DayWeek, Month Specify or Day Specify not blank')
                else:
                    print ('Supplier Delivery Pattern check --- Fail (Please check Week Specify)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2), 'NA', '1,2,3,4,5 separated by comma')

            elif input_to_check == '4':
                month_array = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)).split(',')
                matches_3 = 0
                for month in month_array:
                    if (any(month == x for x in ('B', 'M', 'E'))):
                        matches_3 += 1

                if matches_3 == len(month_array):
                    if str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4)) == '':
                        print ('Supplier Delivery Pattern check --- Pass')
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'Supplier Delivery Pattern filled correctly')
                    else:
                        print ('Supplier Delivery Pattern check --- Fail (Other fields not blank)')
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'DayWeek, Week Specify or Day Specify not blank')
                else:
                    print ('Supplier Delivery Pattern check --- Fail (Please check Month Specify)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3), 'NA', 'B,M,E separated by comma')

            elif input_to_check == '5':
                special_array = str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4)).split(',')
                matches_4 = 0
                for special in special_array:
                    week, day = special.split(':')[0], special.split(':')[1]
                    if (any(week == x for x in ('W1', 'W2', 'W3', 'W4', 'W5'))) and (any(day == x for x in ('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'))):
                        matches_4 += 1

                if matches_4 == len(special_array):
                    if str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)) == '' and str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)) == '':
                        print ('Supplier Delivery Pattern check --- Pass')
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'Supplier Delivery Pattern filled correctly')
                    else:
                        print ('Supplier Delivery Pattern check --- Fail (Other fields not blank)')
                        update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), ', '.join([str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+2)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)), str(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4))]), 'DayWeek, Week Specify or Month Specify not blank')
                else:
                    print ('Supplier Delivery Pattern check --- Fail (Please check Day Specify)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col+4), 'NA', 'Wx:y (x=1,..,5 y=Mon…Sat) separated by comma')

        else:
            print ('Supplier Delivery Pattern check --- Fail (Please check Delivery Pattern %s)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Fixed:1,2,3,4,5')

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
    def customer_contract_details_mod_reference(cell_row):

        # Get concat key
        part_no_customer_contract = str(master_files['xl_sheet_main'].cell_value(cell_row, 3)) + master_files['xl_sheet_main'].cell_value(cell_row, 5)

        # Extract all backup concat into list
        comparison_list_1 = []
        for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
            comparison_list_1.append((row, str(selected['backup_0'].sheet_by_index(0).cell_value(row, 3)) + selected['backup_0'].sheet_by_index(0).cell_value(row, 5)))

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
            for col in range(0, selected['backup_0'].sheet_by_index(0).ncols):
                backup_row_contents.append(selected['backup_0'].sheet_by_index(0).cell_value(backup_row, col))

        submitted_contents = []
        validate_count = 0
        for col in range(0, 25): # Hard Code MAX COLUMN

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

        if validate_count == len(range(2, 25)): # Hard Code MAX COLUMN
            print ('MOD Reference check --- Pass')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', str(submitted_contents), str(backup_row_contents), 'Fields are correctly coloured to indicate \'TO CHANGE\'')

        return True

    def customer_contract_details_customer_part_name(cell_row, cell_col, new_mod):
        if new_mod == 'NEW':
            customer_parts_parts_name = [] # tuple(P/N, Customer Code, Part Name)
            for row in range(9, selected['backup_2'].sheet_by_index(0).nrows):
                customer_parts_parts_name.append((selected['backup_2'].sheet_by_index(0).cell_value(row, 2), selected['backup_2'].sheet_by_index(0).cell_value(row, 3), selected['backup_2'].sheet_by_index(0).cell_value(row, 5)))

            try:
                for row in range(9, additional['TNM_CUSTOMER_PARTS_MASTER'].nrows):
                    if additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 0) == 'NEW':
                        customer_parts_parts_name.append((additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 5)))
                    else:
                        for tuple in customer_parts_parts_name:
                            if additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2) == tuple[0] and additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3) == tuple[1]:
                                customer_parts_parts_name.remove(tuple)
                        customer_parts_parts_name.append((additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 2), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 3), additional['TNM_CUSTOMER_PARTS_MASTER'].cell_value(row, 5)))
            except KeyError:
                pass

            for tuple in customer_parts_parts_name:
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col+1) == tuple[0] and master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3) == tuple[1]:
                    parts_name_ref = tuple[2]

            if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == parts_name_ref:
                print ('Customer Parts Name check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), parts_name_ref, 'Match with Customer Parts Name in Customer Parts Master')
            else:
                print ('Customer Parts Name check --- Warning')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), parts_name_ref, 'Discrepancy with Customer Parts Name in Customer Parts Master, please advise which to register')

        if new_mod == 'MOD':
            print ('Customer Contract Part Name check --- Fail (Cannot be modified in Customer Contract Details)')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Customer Part Name cannot be modified in Customer Contract Details')

    def customer_contract_details_discontinue_mod(cell_row, cell_col, new_mod):
        part_and_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3)

        # Change from 'N' to 'Y' (Discontinue)
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y':
            model_bom_parts_list = []
            for row in range(9, selected['backup_10'].sheet_by_index(0).nrows):
                model_bom_parts_list.append(selected['backup_10'].sheet_by_index(0).cell_value(row, 7) + selected['backup_10'].sheet_by_index(0).cell_value(row, 11))

            if part_and_contract_no not in model_bom_parts_list:
                print ('Discontinue Indicator check --- Pass (Not in Model BOM)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', part_and_contract_no, 'NA', 'P/N not in Model BOM')
            else:
                print ('Discontinue Indicator check --- Fail (Still in Model BOM)')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', part_and_contract_no, 'NA', 'P/N still in Model BOM')

        # Change from 'Y' to 'N' (Revive)
        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N':
            if (any(master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3).find(x) != -1 for x in ('0TBSM', 'TBSJ', 'TBSQ', 'TBSR'))):
                supplier_parts_back = {}
                for row in range(9, selected['backup_3'].sheet_by_index(0).nrows):
                    # supplier_parts_back = [(concat(part no. + supplier code), supplier back no.), ...]
                    supplier_parts_back[selected['backup_3'].sheet_by_index(0).cell_value(row, 2) + selected['backup_3'].sheet_by_index(0).cell_value(row, 3)] = selected['backup_3'].sheet_by_index(0).cell_value(row, 5)

                if (all(supplier_parts_back[master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)] != x for x in ('', 'N/A', 'NA', 'N.A.'))):
                    print ('Revive check (Supplier Back) --- Pass (Require IWRS)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), supplier_parts_back[master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)], 'Revive: Supplier Back No input in Supplier Parts Master')
                else:
                    print ('Revive check (Supplier Back) --- Fail (Check Supplier Back No. in Supplier Parts Master)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), supplier_parts_back[master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col+3)], 'Revive: Supplier Back No not input in Supplier Parts Master')

            submitted_part_customer_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5) + master_files['xl_sheet_main'].cell_value(cell_row, cell_col-4)

            # Append all contracts with same P/N + Customer code and not discontinued
            customer_contract_common = []

            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                if submitted_part_customer_code == str(selected['backup_0'].sheet_by_index(0).cell_value(row, 3)) + str(selected['backup_0'].sheet_by_index(0).cell_value(row, 4)) and selected['backup_0'].sheet_by_index(0).cell_value(row, 8) == 'N':
                    customer_contract_common.append((
                        selected['backup_0'].sheet_by_index(0).cell_value(row, 3),
                        selected['backup_0'].sheet_by_index(0).cell_value(row, 5),
                        selected['backup_0'].sheet_by_index(0).cell_value(row, 8)
                    ))

            if len(customer_contract_common) == 0: # No other non-discontinued contract present
                print('Revive check (Multiple Customer Contract) --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), customer_contract_common, 'Revive: Only 1 Customer Contract No.')
            else:
                # Check if submitted master discontinues those in customer_contract_common
                discontinue_count = 0
                for tuple in customer_contract_common:
                    for row in range(9, master_files['xl_sheet_main'].nrows):
                        if master_files['xl_sheet_main'].cell_value(row, 0) == 'MOD' and master_files['xl_sheet_main'].cell_value(row, 3) == tuple[0] and master_files['xl_sheet_main'].cell_value(row, 5) == tuple[1] and master_files['xl_sheet_main'].cell_value(row, 8) == 'Y':
                            discontinue_count += 1

                if discontinue_count == len(customer_contract_common):
                    print('Revive check (Multiple Customer Contract) --- Pass')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), customer_contract_common, 'Revive: Only 1 Customer Contract No., outstanding parts discontined in submitted master')
                else:
                    print('Revive check (Multiple Customer Contract) --- Fail (Check if other customer contracts have been discontinued)')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), customer_contract_common, 'Revive: Multiple Customer Contract that are not discontinued, no MOD request to discontinue parts')

        # Wrong input
        else:
            print ('Discontinue Indicator check --- Fail (Re-check Input (Y/N))')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'Fail', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Input not Y or N')

    # Check if Customer Contract has been discontinued
    def customer_contract_details_customer_contract_3(cell_row, cell_col, new_mod):
        customer_contract_no = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        backup_customer_contract_no = ()

        for row in range(10, selected['backup_7'].sheet_by_index(0).nrows):
            if selected['backup_7'].sheet_by_index(0).cell_value(row, 2) == customer_contract_no:
                backup_customer_contract_no = (
                    selected['backup_7'].sheet_by_index(0).cell_value(row, 2),
                    selected['backup_7'].sheet_by_index(0).cell_value(row, 16)
                )

        try:
            for row in range(10, additional['TNM_IMP_CUSTOMER_CONTRACT'].nrows):
                if additional['TNM_IMP_CUSTOMER_CONTRACT'].cell_value(row, 2) == customer_contract_no:
                    backup_customer_contract_no = (
                        additional['TNM_IMP_CUSTOMER_CONTRACT'].cell_value(row, 2),
                        additional['TNM_IMP_CUSTOMER_CONTRACT'].cell_value(row, 16)
                    )
        except KeyError:
            pass

        if backup_customer_contract_no == ():
            print('Revive check (Customer Contract Discontinue) --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), customer_contract_no, 'Cannot find customer contract no.')
        elif backup_customer_contract_no[1] == 'N':
            print('Revive check (Customer Contract Discontinue) --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), customer_contract_no, 'Customer Contract is not discontinued')
        elif backup_customer_contract_no[1] == 'Y':
            print('Revive check (Customer Contract Discontinue) --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), customer_contract_no, 'Customer Contract has been discontinued, request user to revive customer contract on screen')
        else:
            print('Revive check (Customer Contract Discontinue) --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), customer_contract_no, 'Strange Error')

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
                if file.find('MRS_TTCContract') != -1:
                    selected['backup_4'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ModuleGroup') != -1:
                    selected['backup_5'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_SupplierContract') != -1:
                    selected['backup_6'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_CustomerContract') != -1 and file.find('MRS_CustomerContractDetail') == -1:
                    selected['backup_7'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ShippingCalendar') != -1:
                    selected['backup_8'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ModelBOM') != -1:
                    selected['backup_10'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
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
                print('WARNING: Not all masters found, program might crash if not all masters available')
                input_to_check = input("Please enter 'Y' to continue, any other key to exit: ")
                if input_to_check != 'Y':
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

        # Checkpoints - PLEASE EDIT
        for row in range(9, master_files['xl_sheet_main'].nrows):

            # Set constant for primary keys per master check
            PRIMARY_KEY_1 = str(master_files['xl_sheet_main'].cell_value(row, 3))
            PRIMARY_KEY_2 = master_files['xl_sheet_main'].cell_value(row, 5)

            # If blank, skip row
            if master_files['xl_sheet_main'].cell_value(row, 2) == '':
                continue

            # Print Part No. header
            print ('%s: Part No. %s' % (str(master_files['xl_sheet_main'].cell_value(row, 0)), str(master_files['xl_sheet_main'].cell_value(row, 3))))
            print()

            # If float, prompt user to change format to String
            if isinstance(master_files['xl_sheet_main'].cell_value(row, 2), float):
                print ('Please change format of Row %d TTC Part No. to Text' % row)
                print ('-' * 10)
                update_df(master_files['xl_sheet_main'].cell_value(row, 0), columns[3], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Please change format of Row ' + str(row + 1) + 'P/N to Text')
                continue

            check_newmod_field(row, 0)

            # Conditional for NEW parts
            if str(master_files['xl_sheet_main'].cell_value(row, 0)).strip(' ') == 'NEW':
                check_maximum_length(row, 'NEW')
                check_compulsory_fields(row, 'NEW')
                customer_contract_details_duplicate_key(row, 3, 'NEW')
                customer_contract_details_part_no_1(row, 3, 'NEW')
                customer_contract_details_part_no_2(row, 3, 'NEW')
                customer_contract_details_customer_contract_1(row, 5, 'NEW')
                customer_contract_details_customer_contract_2(row, 3, 'NEW')
                customer_contract_details_customer_contract_3(row, 5, 'NEW')
                # customer_contract_details_shipping_route(row, 5)
                customer_contract_details_imp(row, 6, 'NEW')
                customer_contract_details_module_group(row, 7, 'NEW')
                customer_contract_details_discontinue_new(row, 8, 'NEW')
                customer_contract_details_ttc_contract_1(row, 9, 'NEW')
                customer_contract_details_ttc_contract_2(row, 9, 'NEW')
                customer_contract_details_ttc_contract_3(row, 9, 'NEW')
                customer_contract_details_ttc_contract_pk(row, 9, 'NEW')
                customer_contract_details_exp(row, 10, 'NEW')
                customer_contract_details_supplier_code(row, 11, 'NEW')
                customer_contract_details_supplier_contract(row, 12, 'NEW')
                customer_contract_details_supplier_delivery_pattern(row, 14, 'NEW')
            # Conditional for MOD parts
            else:
                supplier_delivery_check_cycle = 0
                cols_to_check = get_mod_columns(row)
                if len(cols_to_check) != 0:
                    print('User wishes to MOD the following columns:')
                    # for col in cols_to_check:
                    #     print('%s: %s' % (columns[col+2], master_files['xl_sheet_main'].cell_value(row, col+2)))
                    print()

                    if customer_contract_details_mod_reference(row):

                        check_maximum_length(row, 'MOD')
                        check_compulsory_fields(row, 'MOD')
                        customer_contract_details_duplicate_key(row, 3, 'MOD')
                        customer_contract_details_customer_contract_3(row, 5, 'MOD')

                        # Column specific checks
                        for col in cols_to_check:
                            # Mod: Customer Part Name
                            if col+2 == 2:
                                customer_contract_details_customer_part_name(row, 2, 'MOD')
                            # Mod: TTC Parts No., Customer Code, Customer Contract No.
                            if (any(col+2 == x for x in (3, 4, 5))):
                                print ('%s cannot be modded' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'Cannot be modded')
                            # Mod: Imp Office
                            if col+2 == 6:
                                customer_contract_details_imp(row, 6, 'MOD')
                            # Mod: Module Group Code
                            if col+2 == 7:
                                customer_contract_details_module_group(row, 7, 'MOD')
                                customer_contract_details_ttc_contract_2(row, 9, 'MOD')
                                customer_contract_details_ttc_contract_3(row, 9, 'MOD')
                            # Mod: Discontinue Indicator
                            if col+2 == 8:
                                customer_contract_details_discontinue_mod(row, 8, 'MOD')
                            # Mod: TTC Contract
                            if col+2 == 9:
                                customer_contract_details_ttc_contract_1(row, 9, 'MOD')
                                customer_contract_details_ttc_contract_2(row, 9, 'MOD')
                                customer_contract_details_ttc_contract_3(row, 9, 'MOD')
                                customer_contract_details_ttc_contract_pk(row, 9, 'MOD')
                            # Mod: Exp Office
                            if col+2 == 10:
                                customer_contract_details_exp(row, 10, 'MOD')
                            # Mod: Supplier Code
                            if col+2 == 11:
                                customer_contract_details_supplier_code(row, 11, 'MOD')
                            # Mod: Supplier Contract
                            if col+2 == 12:
                                customer_contract_details_supplier_contract(row, 12, 'MOD')
                            # Mod: Supplier Delivery Pattern
                            if (any(col+2 == x for x in (14, 15, 16, 17, 18))):
                                if supplier_delivery_check_cycle == 0:
                                    customer_contract_details_supplier_delivery_pattern(row, 14, 'MOD')
                                    supplier_delivery_check_cycle += 1
                            # Mod: optional columns
                            if (any(col+2 == x for x in (13, 19, 20, 21, 22, 23, 24))):
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
