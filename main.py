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

# # func to strip non-numeric from SchlCrsID
# def remove_non_num_schlcrsid(list_of_dicts_in):



# func to merg IUID collection w/ class_roster - need pull IUID value from IUID collection, add to ChkDigitInstrctUnitID

# func to find courses missing classroom numbers


if __name__ == '__main__':
    dics = gen_list_of_dicts()
    print(dics)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
