import xlrd, time
import pandas as pd
from os import listdir

# Container Group - Open required workbooks and check against
def container_group():
    from main_workbook import master_files
    from main_workbook import path
    from region_master import region_master

    # Dictionary of columns
    columns = {
        0: "NEW/MOD",
        1: "Reason for the change",
        2: "Container Group Code",
        3: "Description",
        4: "Source Port",
        5: "Destination Port",
        6: "Container Type Code",
        7: "Warehouse Code",
        8: "Discontinue Indicator"
    }

    # Dictionary of required masters for checking
    required = {
        0: "Customer Contract Details Master",
        4: "TTC Contract Master",
        5: "Module Group Master",
        8: "Shipping Calendar Master",
        11: "Container Group Master"
    }

    def check_newmod_field(cell_row, cell_col):
        if all(master_files['xl_sheet_main'].cell_value(cell_row, cell_col) != x for x in ('NEW', 'MOD')):
            print ('NEW/MOD check --- Fail')
            update_df(master_files['xl_sheet_main'].cell_value(cell_row, cell_col), columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Please check NEW/MOD field for whitespace')

    def check_maximum_length(cell_row, new_mod):
        # Hard code range of columns to check
        working_columns = list(range(2, 9))

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
        compulsory_fields = list(range(2, 9))

        if all(master_files['xl_sheet_main'].cell_value(cell_row, col_index) != '' for col_index in compulsory_fields):
            print ('Compulsory Fields check --- Pass')
            update_df(new_mod, 'Compulsory Fields', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', 'NA', 'NA', 'All Compulsory Fields filled')
        else:
            for col_index in compulsory_fields:
                if master_files['xl_sheet_main'].cell_value(cell_row, col_index) == '':
                    print ('Compulsory Fields check --- Fail')
                    update_df(new_mod, columns[col_index], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', columns[col_index] + ' is a Compulsory Field')

    # Check for duplicate primary keys
    def container_group_duplicate_key(cell_row, cell_col, new_mod):
        container_group_list = []
        for row in range(9, master_files['xl_sheet_main'].nrows):
            container_group_list.append((master_files['xl_sheet_main'].cell_value(row, 0), master_files['xl_sheet_main'].cell_value(row, 2)))

        matches = 0
        for container_group_code in container_group_list:
            # Check if part no same modifier
            if master_files['xl_sheet_main'].cell_value(cell_row, 0) == container_group_code[0]:
                if PRIMARY_KEY_1 == container_group_code[1]:
                    matches += 1

        if matches == 1:
            print ('Duplicate Key check --- Pass (Primary key is unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, matches, 'Primary key is unique in submitted master')
        elif matches >1:
            print ('Duplicate Key check --- Fail (Primary key is not unique in submitted master)')
            update_df(new_mod, 'Primary Key(s)', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, matches, 'Primary key is not unique in submitted master')

    # Container Group Code must be unique and in this format: Exp Ctry(2)Imp Ctry(2)-C Number(3)
    def container_group_code_check(cell_row, cell_col, new_mod):
        container_group_list = []
        for row in range(9, selected['backup_11'].sheet_by_index(0).nrows):
            container_group_list.append(selected['backup_11'].sheet_by_index(0).cell_value(row, 2))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) not in container_group_list:
            print ('Container Group Code check 1 --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', PRIMARY_KEY_1, 'NA', 'Container Group Code is unique')
        else:
            backup_row_contents = []
            for row in range(9, selected['backup_11'].sheet_by_index(0).nrows):
                if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == selected['backup_11'].sheet_by_index(0).cell_value(row, 2):
                    for col in range(0, 9): # Hard Code column range
                        backup_row_contents.append(selected['backup_11'].sheet_by_index(0).cell_value(row, col))

            submitted_row_contents = []
            for col in range(0, 9): # Hard Code column range
                submitted_row_contents.append(master_files['xl_sheet_main'].cell_value(cell_row, col))

            discrepancy_reference = []
            for i, cell_value in enumerate(submitted_row_contents):
                if cell_value != backup_row_contents[i] and all(i != x for x in (0, 1, 2)):
                    discrepancy_reference.append(columns[i])

            if len(discrepancy_reference) == 0:
                discrepancy_reference.append('Submitted has no differences from system')

            print ('Container Group Code check 1 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, ', '.join(discrepancy_reference), 'Container Group Code already exists in system')

        code_split = master_files['xl_sheet_main'].cell_value(cell_row, cell_col).split('-')
        # If exp ctry and imp ctry in region master
        if all(x in region_master for x in (code_split[0][:2], code_split[0][2:])):
            if code_split[1][:1] == 'C':
                try:
                    int_transform = int(code_split[1][1:])
                    print ('TTC Contract No. check 2 --- Warning')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', PRIMARY_KEY_1, int_transform, 'TTC Contract No. in correct format, exp/imp ctry found in region master, please check if exp and imp country match other submitted masters')
                except ValueError:
                    print ('TTC Contract No. check 2 --- Fail')
                    update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Exp Ctry(2)Imp Ctry(2)-C Number(3) e.g. THAR-C004')
            else:
                print ('TTC Contract No. check 2 --- Fail')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Exp Ctry(2)Imp Ctry(2)-C Number(3) e.g. THAR-C004')
        else:
            print ('TTC Contract No. check 2 --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', PRIMARY_KEY_1, 'NA', 'Exp Ctry(2)Imp Ctry(2)-C Number(3) e.g. THAR-C004')

    # Warehouse code must be in Warehouse master for Exp Country
    def container_group_warehouse_code(cell_row, cell_col, new_mod):
        warehouse_code = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)

        warehouse_list = []
        for row in range(9, selected['backup_11'].sheet_by_index(0).nrows):
            if selected['backup_11'].sheet_by_index(0).cell_value(row, 7) not in warehouse_list:
                warehouse_list.append(selected['backup_11'].sheet_by_index(0).cell_value(row, 7))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5)[:2] == master_files['xl_sheet_main'].cell_value(cell_row, cell_col)[:2]:
            if warehouse_code in warehouse_list:
                print ('Warehouse Code check --- Pass')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Warehouse Code matches with Exp Country in Container Group Code, found in Container Group Master Backup')
            else:
                print ('Warehouse Code check --- Warning')
                update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'WH Code not found in Container Group Master Backup, please check Warehouse Master on screen')
        else:
            print ('Warehouse Code check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), master_files['xl_sheet_main'].cell_value(cell_row, cell_col-5)[:2], 'Does not match Container Group Code Exp Counry')

    # Source port must match Shipping Calendar Master Exp port
    # Container Group -> Module Group -> CCD -> TTC Contract -> Shipping Calendar
    def container_group_source_port(cell_row, cell_col, new_mod):
        source_port = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        container_group = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-2)

        # Get list of Module Groups to check (system + submitted)
        module_group_list = []
        for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
            if selected['backup_5'].sheet_by_index(0).cell_value(row, 7) == container_group:
                module_group_list.append(selected['backup_5'].sheet_by_index(0).cell_value(row, 2))

        module_group_list = list(set(module_group_list)) # Remove duplicates

        try:
            for row in range(10, additional['TNM_MODULE_GROUP'].nrows):
                if additional['TNM_MODULE_GROUP'].cell_value(row, 7) == container_group and additional['TNM_MODULE_GROUP'].cell_value(row, 0) == 'NEW':
                    module_group_list.append(additional['TNM_MODULE_GROUP'].cell_value(row, 2))
                elif additional['TNM_MODULE_GROUP'].cell_value(row, 0) == 'MOD' and additional['TNM_MODULE_GROUP'].cell_value(row, 2) in module_group_list and additional['TNM_MODULE_GROUP'].cell_value(row, 7) != container_group:
                    module_group_list.remove(additional['TNM_MODULE_GROUP'].cell_value(row, 2))
                elif additional['TNM_MODULE_GROUP'].cell_value(row, 0) == 'MOD' and additional['TNM_MODULE_GROUP'].cell_value(row, 2) not in module_group_list and additional['TNM_MODULE_GROUP'].cell_value(row, 7) == container_group:
                    module_group_list.append(additional['TNM_MODULE_GROUP'].cell_value(row, 2))

        except KeyError:
            pass

        module_group_list = list(set(module_group_list)) # Remove duplicates

        if len(module_group_list) == 0:
            print ('Source Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', source_port, container_group, 'Container Group does not match any Module Group in system')
            return

        # Get list of TTC Contracts to check
        ttc_contract_list = []
        for module_group_code in module_group_list: # for every module group to be checked
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                if selected['backup_0'].sheet_by_index(0).cell_value(row, 7) == module_group_code:
                    ttc_contract_list.append(selected['backup_0'].sheet_by_index(0).cell_value(row, 9))

            ttc_contract_list = list(set(ttc_contract_list)) # Remove duplicates

            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                    if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7) == module_group_code and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'NEW':
                        ttc_contract_list.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9))
                    elif additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'MOD' and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7) == module_group_code and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9) not in ttc_contract_list:
                        ttc_contract_list.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9))
                    elif additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'MOD' and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9) in ttc_contract_list and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7) != module_group_code:
                        ttc_contract_list.remove(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9))

            except KeyError:
                pass

        ttc_contract_list = list(set(ttc_contract_list)) # Remove duplicates

        if len(ttc_contract_list) == 0:
            print ('Source Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', source_port, module_group_list, 'Module Groups of Container Group do not match any Module Group in Customer Contract Details')
            return

        # Get list of Shipping Routes to check
        shipping_route_list = []
        for ttc_contract_no in ttc_contract_list: # for every ttc contract to be checked
            for row in range(9, selected['backup_4'].sheet_by_index(0).nrows):
                if selected['backup_4'].sheet_by_index(0).cell_value(row, 2) == ttc_contract_no:
                    shipping_route_list.append(selected['backup_4'].sheet_by_index(0).cell_value(row, 14))

            shipping_route_list = list(set(shipping_route_list)) # Remove duplicates

            try:
                for row in range(9, additional['TNM_TTC_CONTRACT'].nrows):
                    if additional['TNM_TTC_CONTRACT'].cell_value(row, 2) == ttc_contract_no and additional['TNM_TTC_CONTRACT'].cell_value(row, 0) == 'NEW':
                        shipping_route_list.append(additional['TNM_TTC_CONTRACT'].cell_value(row, 14))
                    elif additional['TNM_TTC_CONTRACT'].cell_value(row, 0) == 'MOD' and additional['TNM_TTC_CONTRACT'].cell_value(row, 14) in shipping_route_list and additional['TNM_TTC_CONTRACT'].cell_value(row, 2) != ttc_contract_no:
                        shipping_route_list.remove(additional['TNM_TTC_CONTRACT'].cell_value(row, 14))

            except KeyError:
                pass

        shipping_route_list = list(set(shipping_route_list)) # Remove duplicates

        if len(shipping_route_list) == 0:
            print ('Source Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', source_port, ttc_contract_list, 'Corresponding TTC Contracts of Module Groups (in CCD) of Container Group do not match any TTC Contract in TTC Contract Master')
            return

        # Get list of Exp Port to check
        port_list = []
        for shipping_route in shipping_route_list: # for every ttc contract to be checked
            for row in range(9, selected['backup_8'].sheet_by_index(0).nrows):
                if selected['backup_8'].sheet_by_index(0).cell_value(row, 2) == shipping_route:
                    port_list.append(selected['backup_8'].sheet_by_index(0).cell_value(row, 6))

        port_list = list(set(port_list)) # Remove duplicates

        if len(port_list) == 0:
            print ('Source Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', source_port, shipping_route_list, 'Corresponding Shipping Routes of TTC Contracts (in TTC Contract Master) of Module Groups (in CCD) of Container Group do not match any Shipping Routes in Shipping Calendar Master')
        elif all(source_port == x for x in port_list):
            print ('Source Port check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', source_port, port_list, 'Source Port matches Shipping Calendar Exp Port')
        else:
            print ('Source Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', source_port, port_list, 'Source Port does not match Shipping Calendar Exp Port')

    # Destination port must match Shipping Calendar Master Imp port
    # Container Group -> Module Group -> CCD -> TTC Contract -> Shipping Calendar
    def container_group_destination_port(cell_row, cell_col, new_mod):
        destination_port = master_files['xl_sheet_main'].cell_value(cell_row, cell_col)
        container_group = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-3)

        # Get list of Module Groups to check (system + submitted)
        module_group_list = []
        for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
            if selected['backup_5'].sheet_by_index(0).cell_value(row, 7) == container_group:
                module_group_list.append(selected['backup_5'].sheet_by_index(0).cell_value(row, 2))

        module_group_list = list(set(module_group_list)) # Remove duplicates

        try:
            for row in range(10, additional['TNM_MODULE_GROUP'].nrows):
                if additional['TNM_MODULE_GROUP'].cell_value(row, 7) == container_group and additional['TNM_MODULE_GROUP'].cell_value(row, 0) == 'NEW':
                    module_group_list.append(additional['TNM_MODULE_GROUP'].cell_value(row, 2))
                elif additional['TNM_MODULE_GROUP'].cell_value(row, 0) == 'MOD' and additional['TNM_MODULE_GROUP'].cell_value(row, 2) in module_group_list and additional['TNM_MODULE_GROUP'].cell_value(row, 7) != container_group:
                    module_group_list.remove(additional['TNM_MODULE_GROUP'].cell_value(row, 2))
                elif additional['TNM_MODULE_GROUP'].cell_value(row, 0) == 'MOD' and additional['TNM_MODULE_GROUP'].cell_value(row, 2) not in module_group_list and additional['TNM_MODULE_GROUP'].cell_value(row, 7) == container_group:
                    module_group_list.append(additional['TNM_MODULE_GROUP'].cell_value(row, 2))

        except KeyError:
            pass

        module_group_list = list(set(module_group_list)) # Remove duplicates

        if len(module_group_list) == 0:
            print ('Destination Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', destination_port, container_group, 'Container Group does not match any Module Group in system')
            return

        # Get list of TTC Contracts to check
        ttc_contract_list = []
        for module_group_code in module_group_list: # for every module group to be checked
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                if selected['backup_0'].sheet_by_index(0).cell_value(row, 7) == module_group_code:
                    ttc_contract_list.append(selected['backup_0'].sheet_by_index(0).cell_value(row, 9))

            ttc_contract_list = list(set(ttc_contract_list)) # Remove duplicates

            try:
                for row in range(9, additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].nrows):
                    if additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7) == module_group_code and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'NEW':
                        ttc_contract_list.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9))
                    elif additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'MOD' and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7) == module_group_code and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9) not in ttc_contract_list:
                        ttc_contract_list.append(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9))
                    elif additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 0) == 'MOD' and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9) in ttc_contract_list and additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 7) != module_group_code:
                        ttc_contract_list.remove(additional['TNM_IMP_CUSTOMER_CONTRACT_DETAI'].cell_value(row, 9))

            except KeyError:
                pass

        ttc_contract_list = list(set(ttc_contract_list)) # Remove duplicates

        if len(ttc_contract_list) == 0:
            print ('Destination Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', destination_port, module_group_list, 'Module Groups of Container Group do not match any Module Group in Customer Contract Details')
            return

        # Get list of Shipping Routes to check
        shipping_route_list = []
        for ttc_contract_no in ttc_contract_list: # for every ttc contract to be checked
            for row in range(9, selected['backup_4'].sheet_by_index(0).nrows):
                if selected['backup_4'].sheet_by_index(0).cell_value(row, 2) == ttc_contract_no:
                    shipping_route_list.append(selected['backup_4'].sheet_by_index(0).cell_value(row, 14))

            shipping_route_list = list(set(shipping_route_list)) # Remove duplicates

            try:
                for row in range(9, additional['TNM_TTC_CONTRACT'].nrows):
                    if additional['TNM_TTC_CONTRACT'].cell_value(row, 2) == ttc_contract_no and additional['TNM_TTC_CONTRACT'].cell_value(row, 0) == 'NEW':
                        shipping_route_list.append(additional['TNM_TTC_CONTRACT'].cell_value(row, 14))
                    elif additional['TNM_TTC_CONTRACT'].cell_value(row, 0) == 'MOD' and additional['TNM_TTC_CONTRACT'].cell_value(row, 14) in shipping_route_list and additional['TNM_TTC_CONTRACT'].cell_value(row, 2) != ttc_contract_no:
                        shipping_route_list.remove(additional['TNM_TTC_CONTRACT'].cell_value(row, 14))

            except KeyError:
                pass

        shipping_route_list = list(set(shipping_route_list)) # Remove duplicates

        if len(shipping_route_list) == 0:
            print ('Destination Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', destination_port, ttc_contract_list, 'Corresponding TTC Contracts of Module Groups (in CCD) of Container Group do not match any TTC Contract in TTC Contract Master')
            return

        # Get list of Exp Port to check
        port_list = []
        for shipping_route in shipping_route_list: # for every ttc contract to be checked
            for row in range(9, selected['backup_8'].sheet_by_index(0).nrows):
                if selected['backup_8'].sheet_by_index(0).cell_value(row, 2) == shipping_route:
                    port_list.append(selected['backup_8'].sheet_by_index(0).cell_value(row, 8))

        port_list = list(set(port_list)) # Remove duplicates

        if len(port_list) == 0:
            print ('Destination Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', destination_port, shipping_route_list, 'Corresponding Shipping Routes of TTC Contracts (in TTC Contract Master) of Module Groups (in CCD) of Container Group do not match any Shipping Routes in Shipping Calendar Master')
        elif all(destination_port == x for x in port_list):
            print ('Destination Port check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', destination_port, port_list, 'Destination Port matches Shipping Calendar Exp Port')
        else:
            print ('Destination Port check --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', destination_port, port_list, 'Destination Port does not match Shipping Calendar Exp Port')

    # Container Type Code must be found in Container Type Master
    def container_group_container_type(cell_row, cell_col, new_mod):
        container_type_list = []
        for row in range(9, selected['backup_11'].sheet_by_index(0).nrows):
            container_type_list.append(selected['backup_11'].sheet_by_index(0).cell_value(row, 6))

        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) in container_type_list:
            print ('Container Type check --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Container Type registered before in Container Master')
        else:
            print ('Container Type check --- Warning')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'WARNING', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Container Type not registered before, please check if found in Container Type Master')

    # All related Module Groups must be discontinued
    def container_group_discontinue(cell_row, cell_col, new_mod):
        if master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'Y' and new_mod == 'NEW':
            print ('Discontinue Indicator check NEW --- Fail')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'NEW Container Group should not be registered as discontinued')
        elif master_files['xl_sheet_main'].cell_value(cell_row, cell_col) == 'N' and new_mod == 'NEW':
            print ('Discontinue Indicator check NEW --- Pass')
            update_df(new_mod, columns[cell_col], cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'PASS', master_files['xl_sheet_main'].cell_value(cell_row, cell_col), 'NA', 'Discontinue Indicator valid')
        elif new_mod == 'MOD':
            # Check if all non-discontinued parts in CCD for Module Group No. are being MOD to discontinue in submitted CCD
            # CCD tuple = (part no., customer contract no., module group no., discontinue indicator)
            containger_group = master_files['xl_sheet_main'].cell_value(cell_row, cell_col-6)
            submitted_module_group = []
            for row in range(10, selected['backup_5'].sheet_by_index(0).nrows):
                if selected['backup_5'].sheet_by_index(0).cell_value(row, 7) == container_group:
                    submitted_module_group.append(selected['backup_5'].sheet_by_index(0).cell_value(row, 2))

            try:
                for row in range(10, additiona['TNM_MODULE_GROUP'].nrows):
                    if additiona['TNM_MODULE_GROUP'].cell_value(row, 7) == container_group:
                        submitted_module_group.append(additiona['TNM_MODULE_GROUP'].cell_value(row, 2))
            except KeyError:
                pass

            submitted_module_group = list(set(submitted_module_group))

            customer_contract_details = []
            for row in range(9, selected['backup_0'].sheet_by_index(0).nrows):
                customer_contract_details.append((selected['backup_0'].sheet_by_index(0).cell_value(row, 3), selected['backup_0'].sheet_by_index(0).cell_value(row, 5), selected['backup_0'].sheet_by_index(0).cell_value(row, 7), selected['backup_0'].sheet_by_index(0).cell_value(row, 8)))

            # CCD discontinue = (primary key concat, discontinue indicator)
            contract_no_to_discontinue = []
            for tuple in customer_contract_details:
                for code in submitted_module_group:
                    if tuple[2] == code and tuple[3] == 'N':
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
    def container_group_mod_reference(cell_row):

        # Get concat key
        container_group_no = str(master_files['xl_sheet_main'].cell_value(cell_row, 2))

        # Extract all backup concat into list
        comparison_list_1 = []
        for row in range(9, selected['backup_11'].sheet_by_index(0).nrows):
            comparison_list_1.append((row, str(selected['backup_11'].sheet_by_index(0).cell_value(row, 2))))

        # Find backup row
        backup_row = 0
        for concat_str in comparison_list_1:
            if container_group_no == concat_str[1]:
                backup_row = concat_str[0]

        # If cannot find, return False
        if backup_row == 0:
            print ('MOD Reference check --- Fail')
            update_df('MOD', 'ALL', cell_row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', 'NA', 'NA', 'Cannot find Backup Row in system for MOD part')
            return False

        if backup_row >= 0:
            # Extract contents of backup row
            backup_row_contents = []
            for col in range(0, selected['backup_11'].sheet_by_index(0).ncols):
                backup_row_contents.append(selected['backup_11'].sheet_by_index(0).cell_value(backup_row, col))

        submitted_contents = []
        validate_count = 0
        for col in range(0, 9): # Hard Code MAX COLUMN

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

        if validate_count == len(range(2, 9)): # Hard Code MAX COLUMN
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
                if file.find('MRS_TTCContract') != -1:
                    selected['backup_4'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ModuleGroup') != -1:
                    selected['backup_5'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ShippingCalendar') != -1:
                    selected['backup_8'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
                    print ('Retrieved file: %s' % file)
                if file.find('MRS_ContainerGroup') != -1:
                    selected['backup_11'] = xlrd.open_workbook(path + '\\2) Backup\\' + file)
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

            # If blank, skip row
            if master_files['xl_sheet_main'].cell_value(row, 2) == '':
                continue

            # Print Container Group No. header
            print ('%s: Part No. %s' % (str(master_files['xl_sheet_main'].cell_value(row, 0)), str(master_files['xl_sheet_main'].cell_value(row, 2))))
            print()

            check_newmod_field(row, 0)

            # Conditional for NEW parts
            if str(master_files['xl_sheet_main'].cell_value(row, 0)).strip(' ') == 'NEW':
                check_maximum_length(row, 'NEW')
                check_compulsory_fields(row, 'NEW')
                container_group_duplicate_key(row, 2, 'NEW')
                container_group_code_check(row, 2, 'NEW')
                container_group_warehouse_code(row, 7, 'NEW')
                container_group_source_port(row, 4, 'NEW')
                container_group_destination_port(row, 5, 'NEW')
                container_group_container_type(row, 6, 'NEW')
                container_group_discontinue(row, 8, 'NEW')

            # Conditional for MOD parts
            else:
                cols_to_check = get_mod_columns(row)
                if len(cols_to_check) != 0:
                    print('User wishes to MOD the following columns:')
                    for col in cols_to_check:
                        print('%s: %s' % (columns[col+2], master_files['xl_sheet_main'].cell_value(row, col+2)))
                    print()

                    if container_group_mod_reference(row):
                        check_maximum_length(row, 'MOD')
                        check_compulsory_fields(row, 'MOD')
                        container_group_duplicate_key(row, 2, 'MOD')

                        for col in cols_to_check:
                            # Mod: Container Group Code
                            if col+2 == 2:
                                print ('%s cannot be modded' % columns[col+2])
                                update_df('MOD', columns[col+2], row, PRIMARY_KEY_1, PRIMARY_KEY_2, 'FAIL', master_files['xl_sheet_main'].cell_value(row, col+2), 'NA', 'Cannot be modded')
                            # Mod: Source Port
                            if col+2 == 4:
                                container_group_source_port(row, 4, 'MOD')
                            # Mod: Destination Port
                            if col+2 == 5:
                                container_group_destination_port(row, 5, 'MOD')
                            # Mod: Container Type
                            if col+2 == 6:
                                container_group_container_type(row, 6, 'MOD')
                            # Mod: Warehouse Code
                            if col+2 == 7:
                                container_group_warehouse_code(row, 7, 'MOD')
                            # Mod: Discontinue Indicator
                            if col+2 == 8:
                                container_group_discontinue(row, 8, 'MOD')
                            # Mod: optional columns
                            if col+2 == 3:
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
