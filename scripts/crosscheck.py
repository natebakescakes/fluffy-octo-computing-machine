import os

import xlrd

def crosscheck(submitted_wb, temp_dir):
    """
    Checks whether Submitted MRS and Temp Files match in value
    Will not check MRS when it is split into multiple Temp Files

    submitted_wb: xlrd workbook object, submitted MRS
    temp_dir: string, directory of "5) Temp"
    """
    submitted_sheets = sorted(submitted_wb.sheets(), key=lambda x: x.name)

    temp_sheets = sorted([xlrd.open_workbook(os.path.join(temp_dir, temp_wb_path)).sheet_by_index(0) \
        for temp_wb_path in os.listdir(temp_dir)], key=lambda x: x.name)

    # Remove duplicates
    sheet_name_list = [sheet.name for sheet in temp_sheets]
    for sheet_name in sheet_name_list:
        if sheet_name_list.count(sheet_name) > 1:
            print ('More than 1 {}'.format(sheet_name))
            sheet_name_list.remove(sheet_name)

            for sheet in submitted_sheets:
                if sheet.name == sheet_name:
                    submitted_sheets.remove(sheet)

            for sheet in temp_sheets:
                if sheet.name == sheet_name:
                    temp_sheets.remove(sheet)

    if len(temp_sheets) == 0:
        print ('No Temp Sheets')
        return False
    elif len(submitted_sheets) == 0:
        print ('No Submitted Sheets')
        return False

    for temp_sheet, submitted_sheet in zip(temp_sheets, submitted_sheets):
        if all(temp_sheet.cell_value(row, col) == submitted_sheet.cell_value(row, col) \
            for row in range(9, submitted_sheet.nrows) \
            for col in range(submitted_sheet.ncols)):
            return True
        else:
            for row in range(9, submitted_sheet.nrows):
                for col in range(submitted_sheet.ncols):
                    if temp_sheet.cell_value(row, col) != submitted_sheet.cell_value(row, col):
                        print (submitted_sheet.name,
                            row+1,
                            col+1,
                            submitted_sheet.cell_value(row, col),
                            temp_sheet.cell_value(row, col)
                        )

            return False
