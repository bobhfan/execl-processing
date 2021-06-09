import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import PatternFill, Alignment
from datetime import datetime, timedelta
from datetime import date as date_op
import pprint
import argparse
import json
from collections import OrderedDict, defaultdict
import holidays
import sys
import os

def get_xlsx_file_list(dir):
    file_list = []
    data_files = [(x[0], x[2]) for x in os.walk(dir)]
    print(data_files)
    for path_files in data_files:
        for file_name in path_files[1]:
            if '.xlsx' in file_name and 'Zone' not in file_name:
                os.rename(dir+file_name, dir+file_name.lower())
                file_list.append(file_name.lower())
    return set(file_list)


def read_forcast_json(file_name):

    # FMT = '%d-%b-%Y'
    FMT = '%Y-%m-%d'
    YEAR_FMT = '%Y'

    if '7-11' in file_name:
        print('file with error')

    current_date = datetime.now().date()
    previuos_year_date = date_op(current_date.year - 1, current_date.month, current_date.day)
    next_year_date = date_op(current_date.year + 1, current_date.month, current_date.day)

    print('now is {} pre is {} next is {}'.format(current_date, previuos_year_date, next_year_date))

    pre_year = '2020'
    cur_year = '2021'

    # date = datetime.now()


    wb= openpyxl.load_workbook(file_name, data_only=True)
    print(wb.sheetnames)
    sheet = wb.active

    # for line in range(1, sheet.max_row+1):
    # for line in range(3, 4):
    #     for colum_n in range(1, sheet.max_column+1):
    #         print("{} and {}".format(sheet.cell(row=line, column=colum_n).value, type(sheet.cell(row=line, column=colum_n).value)))

    start_row = 3
    field_name_part_1_dict = {}
    field_name_dict = {}

    customizd_max_column = 73
    first_round = True

    for x in range (start_row,start_row + 1):
        for y in range(1, customizd_max_column):
            # print("field {} at {}:{}".format(sheet.cell(row=x,column=y).value, start_row, y))
            append_str = ''
            if sheet.cell(row=x,column=y).value:
                field_name = sheet.cell(row=x,column=y).value
            if 'JAN' == field_name:
                first_round = False
            if y > 5:
                if first_round:
                    append_str = pre_year
                else:
                    append_str = cur_year
            if append_str:
                field_output_name = append_str + '\n' + field_name
            else:
                field_output_name = field_name
            field_name_part_1_dict.update({y:field_output_name})



    for x in range (start_row + 1, start_row + 2 ):
        for y in range(1, customizd_max_column):
            if sheet.cell(row=x,column=y).value:
                part_1 = field_name_part_1_dict[y]
                field_name_dict[y] = part_1 + '\n' + sheet.cell(row=x,column=y).value
            else:
                field_name_dict[y] = field_name_part_1_dict[y]

    pprint.pprint(field_name_dict)

    print("Ther are {} rows in {} sheet of {} ".format(sheet.max_row, wb.sheetnames[0], file_name ))
    seq_no = 1
    data_dict = {}
    for line in range(start_row + 2, sheet.max_row+1):
    # for line in range(7, 8):
        tmp_dict = {}
        item_no = ''
        for colum_n in range(1, customizd_max_column):
            # This is for extract item list info to create the dict for the first step
            # {item_no : { interesting_field : value}}
            if field_name_dict[colum_n]:
                if 'Item &' == field_name_dict[colum_n] or 'Item No.' == field_name_dict[colum_n]:
                    if isinstance(sheet.cell(row=line, column=colum_n).value, float):
                        item_no = str(int(sheet.cell(row=line, column=colum_n).value))
                    else:
                        item_no = sheet.cell(row=line, column=colum_n).value

                else:
                    if isinstance(sheet.cell(row=line, column=colum_n).value, datetime):
                        date_value = sheet.cell(row=line, column=colum_n).value.strftime(FMT)
                        tmp_dict.update({field_name_dict[colum_n] : date_value})
                    else:
                        tmp_dict.update({field_name_dict[colum_n] : sheet.cell(row=line, column=colum_n).value})
        if item_no:
            data_dict[item_no] = tmp_dict

    with open(file_name.replace('.xlsx', '_step_1.json'), "w") as json_file:
        json.dump(data_dict, json_file, indent = 4)

def read_forcast(dir):
    files = get_xlsx_file_list(dir)
    for file in files:
        read_forcast_json(dir+file)