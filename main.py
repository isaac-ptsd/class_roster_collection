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
course_roster_worksheet = sh.sheet1


def gen_list_of_dicts(worksheet_in):
    # pulling all data from the spreadsheet with one API call
    return worksheet_in.get_all_records()  # spreadsheet data saved as a list of dictionaries


def search(iuid_list, school, section, course):
    for element in iuid_list:
        if element["school"] == school and element["section"] == section and element["course"] == course:
            return element["iuid"]


def alpha_stripper(string_in):
    # returns a string with all alpha characters stripped
    return ''.join(c for c in str(string_in) if c.isdigit())


def remove_alphas_schlcrsid(list_of_dicts_in):
    cell_list = course_roster_worksheet.range('E2:E' + str(course_roster_worksheet.row_count))
    stripped_SchlCrsID = [alpha_stripper(student["SchlCrsID"]) for student in list_of_dicts_in]
    for i, val in enumerate(stripped_SchlCrsID):
        cell_list[i].value = val
    course_roster_worksheet.update_cells(cell_list)
    print("All alpha chars removed from SchlCrsID")


# func to merg IUID collection w/ class_roster - need pull IUID value from IUID collection, add to ChkDigitInstrctUnitID
def merge_iuid_w_class_roster(in_sheet_key_iuid, list_of_dicts_in):

    # iuid_sect_course_school = {
    #   "iuid": 123456
    #   "school": 370
    #   "section": "100",
    #   "course":  cr101,
    # }

    iuid_sh: Spreadsheet = gc.open_by_key(in_sheet_key_iuid)
    iuid_worksheet = iuid_sh.sheet1
    iuid_school_sect_course = []
    iuid_dicts = gen_list_of_dicts(iuid_worksheet)

    for element in iuid_dicts:
        iuid_school_sect_course.append({
            'iuid': element["ChkDigitInstrctUnitID"],
            'school': element["SchlInstID"],
            'section': element["SchlSectID"],
            'course': element["SchlCrsID"]
        })

    cell_list = course_roster_worksheet.range('A2:A' + str(course_roster_worksheet.row_count))

    i = 0
    for record in list_of_dicts_in:
        # iterate through course_record sheet data and find the corresponding IUID from the IUID_collection data
        # using the search() function, building a cell_list of IUID's to be writen to course_recrd sheet
        cell_list[i].value = search(iuid_school_sect_course, record["SchlInstID"], record["SchlSectID"], record["SchlCrsID"])
        i += 1
    course_roster_worksheet.update_cells(cell_list)


# func to search course_roster sheet for rows missing IUID's, then filter down to unique: schoolid, section, course
def find_missing_iuid(cr_list_of_dicts_in):
    missing_iuid_list = [row for row in cr_list_of_dicts_in if row["ChkDigitInstrctUnitID"] == '']
    # filter to unique: schoolid, section, course
    school_section_course = set([(row["SchlInstID"], row["SchlSectID"], row["CrsCd"]) for row in missing_iuid_list])
    print(school_section_course)


# func to find courses missing classroom numbers

if __name__ == '__main__':
    cr_dicts = gen_list_of_dicts(course_roster_worksheet)

    # remove_alphas_schlcrsid(cr_dicts)
    # merge_iuid_w_class_roster("1fR2e7oLFPRAO1Re9oiUTRvJJid8UmJjzqY5NJSs3ELw", cr_dicts)
    find_missing_iuid(cr_dicts)
