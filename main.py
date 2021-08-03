import gspread
from gspread import Spreadsheet
from gspread import utils
from gspread_dataframe import set_with_dataframe
from datetime import datetime as dt
from pandas import DataFrame


sheet_key = '1N-4B26kQSS3eUcwp_Kw3Ues0KUj_tYX4EQahGSdbtCA'  # class roster dev sheet
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


# TODO: can't just strip alphas need to use a case statement or something; K1 -> KG,
# also can't strip alphas from PHS courses
def remove_alphas_schlcrsid(list_of_dicts_in):
    # note: if course == 'K1', change to 'KG'
    cell_list = course_roster_worksheet.range('E2:E' + str(course_roster_worksheet.row_count))
    stripped_SchlCrsID = [alpha_stripper(student["SchlCrsID"]) for student in list_of_dicts_in]
    for student in list_of_dicts_in:
        if student["SchlCrsID"] == 'k1':
            student["SchlCrsID"] = 'KG'
        if student["SchlCrsID"] == 'e1':
            student["SchlCrsID"] = '1'
        if student["SchlCrsID"] == 'e2':
            student["SchlCrsID"] = '2'
        if student["SchlCrsID"] == 'e3':
            student["SchlCrsID"] = '3'
        if student["SchlCrsID"] == 'e4':
            student["SchlCrsID"] = '4'
        if student["SchlCrsID"] == 'e5':
            student["SchlCrsID"] = '5'
    for i, val in enumerate(stripped_SchlCrsID):
        cell_list[i].value = val
    course_roster_worksheet.update_cells(cell_list)
    print("All alpha chars removed from SchlCrsID")


# func to merg IUID collection w/ class_roster - get IUID value from IUID collection, add to ChkDigitInstrctUnitID
# todo: use CrsCd NOT SchlCrsID
def merge_iuid_w_class_roster(in_sheet_key_iuid, list_of_dicts_in):
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
        cell_list[i].value = \
            search(iuid_school_sect_course, record["SchlInstID"], record["SchlSectID"], record["SchlCrsID"])
        i += 1
    course_roster_worksheet.update_cells(cell_list)


# func to search course_roster sheet for rows missing IUID's, then filter down to unique: schoolid, section, course
def find_missing_iuid(cr_list_of_dicts_in):
    missing_iuid_list = [row for row in cr_list_of_dicts_in if row["ChkDigitInstrctUnitID"] == '']
    # filter to unique: schoolid, section, course
    return set([(row["SchlInstID"], row["SchlSectID"], row["SchlCrsID"]) for row in missing_iuid_list])


# func to find courses missing classroom numbers
def find_courses_missing_classnum(list_of_dicts_in):
    return [ele for ele in list_of_dicts_in if ele["ClsRmID"]]


def add_wsheet(data_in, sheet_name, email_in='isaac.stoutenburgh@phoenix.k12.or.us'):
    """
    :param data_in: List of dictionaries
    :param sheet_name: String
    :param email_in: String: defaults to 'isaac.stoutenburgh@phoenix.k12.or.us'
    :return: No return value
             Will add a new worksheet to the spreadsheet
    """
    if not data_in:
        print("add_wsheet: data_in is empty; will not attempt to add to worksheet")
    else:
        try:
            if data_in[0]:
                headers = list(data_in[0].keys())
            else:
                headers = list(data_in.keys())
            # +1 fixes bug when data_in has only one record
            sheet = sh.add_worksheet(sheet_name, len(data_in) + 1, len(headers))
            sheet.append_row(headers)
            last_cell = gspread.utils.rowcol_to_a1(len(data_in), len(headers))
            cell_range = sheet.range('A2:' + last_cell)
            flattened_test_data = []
            for row in data_in:
                for column in headers:
                    flattened_test_data.append(row[column])
            for i, cell in enumerate(cell_range):
                cell.value = flattened_test_data[i]
            sheet.update_cells(cell_range)
        except TypeError as e:
            print("Worksheet not created - no data", e)
        except IndexError as e:
            print("\nERROR in function: add_wsheet ", e)
        except gspread.exceptions.APIError as e:
            print("ERROR ADDING WORKSHEET: ", e)


def mmddyyyy_to_dt_obj(string_in):
    chars = [char for char in string_in]
    if len(chars) == 7:
        date_str = f"0{chars[0]}/{chars[1]}{chars[2]}/{chars[3]}{chars[4]}{chars[5]}{chars[6]}"
    if len(chars) == 8:
        date_str = f"{chars[0]}{chars[1]}/{chars[2]}{chars[3]}/{chars[4]}{chars[5]}{chars[6]}{chars[7]}"
    return dt.strptime(date_str, '%m/%d/%Y').date()


def add_sub(cr_list_of_dicts_in):
    #  need to know:
    #  * dates sub taught for
    #  * full time teacher the sub covered for
    #
    sub_name = input("Enter the substitutes name: ")
    sub_id = input("Enter the substitutes staff id: ")
    sub_gender = input("Enter the substitutes gender: ")
    sub_start_date = input("Enter the date the sub started (mmddyyyy): ")
    sub_end_date = input("Enter the date the sub ended (mmddyyyy): ")
    teacher_id = input("Enter the teachers staff id of the teacher that the sub is covering for: ")

    print('you entered: ', sub_name, sub_id, sub_gender, sub_start_date, sub_end_date, teacher_id)
    # look up teacher_id
    # duplicate rows (x2)
    # for one set of duplicates replace:
    # * teacher name
    # * teacher id
    # * teacher gender
    # fix dates
    # * og teacher rows need end date = sub start date
    # * new teacher rows need start date = sub end date
    # all student start end dates must fall within teacher/sub start/end dates

    found_tch_list = [ele for ele in cr_list_of_dicts_in if ele["EmplyrStaffID"] == teacher_id]
    found_tch_df = DataFrame(found_tch_list)

    # update with subs info:


    print(found_tch_list)
    print(found_tch_df)

    # logic paths:
    # 1) Triple teachers rows if start date and end date more than 10 days from start/end of course
    #   * update dates and fill in subs info where appropriate
    # 2) Duplicate rows if start date or end date within 10 days of start/end of course dates
    #   * update dates and fill in subs info where appropriate

    set_with_dataframe(course_roster_worksheet, found_tch_df, row=len(cr_list_of_dicts_in)+2, col=1, include_column_header=False)
    date1 = str(found_tch_list[0]["CrsBeginDtTxt"])

    print(mmddyyyy_to_dt_obj(date1) - mmddyyyy_to_dt_obj(sub_start_date))




if __name__ == '__main__':
    cr_dicts = gen_list_of_dicts(course_roster_worksheet)

    # merge_iuid_w_class_roster("1fR2e7oLFPRAO1Re9oiUTRvJJid8UmJjzqY5NJSs3ELw", cr_dicts)
    # print(find_missing_iuid(cr_dicts))
    # print(len(find_missing_iuid(cr_dicts)))
    # print(find_courses_missing_classnum(cr_dicts))
    # add_wsheet(find_courses_missing_classnum(cr_dicts), "courses missing rooms")
    add_sub(cr_dicts)

