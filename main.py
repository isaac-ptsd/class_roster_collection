import gspread
from gspread import Spreadsheet
from gspread import utils
import csv
import time
from datetime import datetime as dt


sheet_key = '1UM5mDB_FAgJc_LFpVqi7KID9YP1EBTEx_cYD8N12H1I'  # class roster
# authorize, and open a google spreadsheet
gc = gspread.oauth()
sh: Spreadsheet = gc.open_by_key(sheet_key)
worksheet = sh.sheet1


def gen_list_of_dicts():
    # pulling all data from the spreadsheet with one API call
    return worksheet.get_all_records()  # spreadsheet data saved as a list of dictionaries


def alpha_stripper(string_in):
    # returns a string with all alpha characters stripped
    return ''.join(c for c in str(string_in) if c.isdigit())


def remove_alphas_schlcrsid(list_of_dicts_in):
    cell_list = worksheet.range('E2:E' + str(worksheet.row_count))
    stripped_SchlCrsID = [alpha_stripper(student["SchlCrsID"]) for student in list_of_dicts_in]
    for i, val in enumerate(stripped_SchlCrsID):
        cell_list[i].value = val
    worksheet.update_cells(cell_list)
    print("All alpha chars removed from SchlCrsID")


# func to merg IUID collection w/ class_roster - need pull IUID value from IUID collection, add to ChkDigitInstrctUnitID
def merge_iuid_w_class_roster(in_sheet_key_iuid):

    # iuid_sect_course_school = {
    #   "iuid": 123456
    #   "school": 370
    #   "section": "100",
    #   "course":  cr101,
    # }

    iuid_sh: Spreadsheet = gc.open_by_key(in_sheet_key_iuid)
    iuid_worksheet = iuid_sh.sheet1
    iuid_school_sect_course = [{}]
    for row in iuid_worksheet:
        iuid_school_sect_course.append({
            'iuid': row["ChkDigitInstrctUnitID"],
            'school': row["SchlInstID"],
            'section': row["SchlSectID"],
            'course': row["SchlCrsID"]
        })




# func to find courses missing classroom numbers


if __name__ == '__main__':
    dicts = gen_list_of_dicts()
    remove_alphas_schlcrsid(dicts)


