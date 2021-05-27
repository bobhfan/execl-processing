import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta
from datetime import date as date_op
import pprint
import argparse
import json
from collections import OrderedDict, defaultdict
import holidays
import sys

FMT = '%Y-%m-%d'
YEAR_FMT = '%Y'
  

MAXIMUM_DUPLICATE_ORDER = 100

# time = datetime.strptime("2021-02-13 00:00:00", FMT)

# date_time = datetime.now()

# print("now is {}".format(date_time.strftime(FMT)))

WEEK_DAY = {
    0 : 'Mon',
    1 : "Tue",
    2 : "Wen",
    3 : "Thu",
    4 : "Fri",
    5 : "Sat",
    6 : "Sun",
}

WORKING_DIR = ''

file_dict = {
                WORKING_DIR + 'BOM list.xlsx' : 1,
                WORKING_DIR + 'Item List.xlsx' : 1,
                WORKING_DIR + 'customer order list.xlsx': 1,
                WORKING_DIR + 'Purchase Lines.xlsx':1,
            }


Item_List_Interest_Field_List = ['Duty Class', 
                                'W1',
                                'W2',
                                'Base Unit of Measure' ,
                                'Vendor No.',
                                'Item Category Code',
                                'Quantity on Hand',
                                ]

Interest_Duty_List = ["FILM", "BAG",  "SEASONING"]

Valid_Purchase_Period = timedelta(days=365)
Current_day = datetime.today() + timedelta(days=-1)

Purchase_Lines_Interest_Dict = {
        'Promised Receipt Date':'Promised Receipt Date',
        'Quantity' : 'Purchase Quantity',
        'Requested Receipt Date' : 'Requested Receipt Date',
        'Outstanding Quantity' : 'Outstanding Quantity' ,
    }

Dict_step_2 = {
                WORKING_DIR + 'BOM list_step_1.json' : {},
                WORKING_DIR + 'Item List_step_1.json' : {},
                WORKING_DIR + 'customer order list_step_1.json': {},
                WORKING_DIR + 'Purchase Lines_step_1.json' : {},
                }

def read_data_from_files(file_dict):

    for file_name, start_row in file_dict.items():
        wb= openpyxl.load_workbook(file_name)
        names = wb.sheetnames
        print(wb.sheetnames)
        sheet = wb.active

        data_dict = {}

        # key -> column_no
        # value ->header description
        field_name_dict = {}

        for x in range (start_row,start_row + 1):

            for y in range(1,sheet.max_column + 1):
                field_name_dict.update({y:sheet.cell(row=x,column=y).value})
        
        pp = pprint.PrettyPrinter(indent=4)
        pp.pprint(field_name_dict)

        print("{}'s total row is {}".format(file_name, sheet.max_row))

        # iterate through excel
        seq_no = 1
        for line in range(start_row + 1, sheet.max_row+1):
        # for line in range(3, 155):
            tmp_dict = {}
            for colum_n in range(1, sheet.max_column+1):
                
                # This is for extract item list info to create the dict for the first step
                # {item_no : { interesting_field : value}}
                if 'Item List.xlsx' in file_name :
                    if 'No.' in field_name_dict[colum_n] \
                        and ' ' not in field_name_dict[colum_n]:

                        if isinstance(sheet.cell(row=line, column=colum_n).value, float):
                            item_no = str(int(sheet.cell(row=line, column=colum_n).value))
                        else:
                            item_no = sheet.cell(row=line, column=colum_n).value

                    elif field_name_dict[colum_n] in Item_List_Interest_Field_List:
                        tmp_dict.update({field_name_dict[colum_n] : sheet.cell(row=line, column=colum_n).value}) 
                    
                    elif 'Description' in field_name_dict[colum_n] and ' ' not in field_name_dict[colum_n]:
                        tmp_dict.update({field_name_dict[colum_n] : sheet.cell(row=line, column=colum_n).value})                    

                    elif 'Blocked' in field_name_dict[colum_n]:
                        item_block = sheet.cell(row=line, column=colum_n).value

                # This is for extract purchase line info to create the dict for the first step
                # {item_no : {po_no : { Interest_description : value }}}
                if 'Purchase Lines.xlsx' in file_name :
                    if 'No.' in field_name_dict[colum_n] and ' ' not in field_name_dict[colum_n] :
                        item_no = sheet.cell(row=line, column=colum_n).value

                    elif 'Document No.' in field_name_dict[colum_n] :
                        document_no = sheet.cell(row=line, column=colum_n).value

                    elif "Promised Receipt Date" in field_name_dict[colum_n] :
                        if sheet.cell(row=line, column=colum_n).value:
                            promise_date = sheet.cell(row=line, column=colum_n).value.strftime(FMT)
                    
                    elif "Outstanding Quantity" in field_name_dict[colum_n] :
                        quantity = sheet.cell(row=line, column=colum_n).value

                    elif "Requested Receipt Date" in field_name_dict[colum_n] :
                        if sheet.cell(row=line, column=colum_n).value and isinstance(sheet.cell(row=line, column=colum_n).value, datetime):
                            request_date = sheet.cell(row=line, column=colum_n).value.strftime(FMT)

                # {bom_no : {item_no : { production_variable : value}}}
                # To have a specific production done may need multiple items involved in the production
                elif 'BOM list.xlsx' in file_name :
                    
                    if 'Production BOM No.' in field_name_dict[colum_n]:
                        dict_key = sheet.cell(row=line, column=colum_n).value 
                        # have this code first time should create empty dict for further
  

                    elif 'No.' in field_name_dict[colum_n] and ' ' not in field_name_dict[colum_n]:
                        material_dict_key = sheet.cell(row=line, column=colum_n).value 

                    else:
                        if "Description" not in field_name_dict[colum_n]:
                            if isinstance(sheet.cell(row=line, column=colum_n).value, datetime):
                                tmp_dict.update(
                                    {field_name_dict[colum_n] : 
                                    sheet.cell(row=line, column=colum_n).value.strftime(FMT)})
                            else:

                                tmp_dict.update({
                                    field_name_dict[colum_n] : 
                                    sheet.cell(row=line, column=colum_n).value})     
                
                # each customer order associated with a specific bom_no
                # {customer_order+bom_no : {customer_order_field : value}}
                elif 'customer order list.xlsx' in file_name :
                    if 'No.' in field_name_dict[colum_n] \
                        and ' ' not in field_name_dict[colum_n]:
                        dict_key = sheet.cell(row=line, column=colum_n).value     

                    elif 'Document No.' in field_name_dict[colum_n]:
                        customer_document_no = sheet.cell(row=line, column=colum_n).value  

                    elif 'Outstanding Quantity' in field_name_dict[colum_n] and not ' ' in field_name_dict[colum_n]:
                        tmp_dict.update({
                                'Outstanding Quantity' : 
                                sheet.cell(row=line, column=colum_n).value})      

                    else:
                        if "Description" not in field_name_dict[colum_n]:
                            if isinstance(sheet.cell(row=line, column=colum_n).value, datetime):
                                tmp_dict.update(
                                    {field_name_dict[colum_n] : 
                                    sheet.cell(row=line, column=colum_n).value.strftime(FMT)})
                            else:

                                tmp_dict.update({
                                    field_name_dict[colum_n] : 
                                    sheet.cell(row=line, column=colum_n).value})   

            
            # when the processing end for one row should update the "No.->":Duty Class into the dict
            if 'Item List.xlsx' in file_name :
                # if not item_block or 'No' in item_block :
                data_dict.update({item_no : tmp_dict})

            elif 'customer order list.xlsx' in file_name:
                real_key = customer_document_no + '---' + str(dict_key)
                data_dict.update({real_key: tmp_dict})

            elif 'BOM list.xlsx' in file_name:
                if not material_dict_key:
                    material_dict_key = 'unknown'
                    print("BOM list err: empty BOM No. for {} ".format(dict_key))
                if dict_key not in data_dict:
                    data_dict.update({dict_key:{}})

                data_dict[dict_key].update({
                            material_dict_key : 
                            tmp_dict})

            elif 'Purchase Lines.xlsx' in file_name:
                if item_no not in data_dict:
                    data_dict[item_no] = {}
                    data_dict[item_no]['promise'] = {}
                    data_dict[item_no]['request'] = {}

                if 'promise_date' in locals():
                    data_dict[item_no]['promise'][promise_date + '---' + str(seq_no)] = quantity
                elif 'request_date' in locals():
                    data_dict[item_no]['request'][request_date + '---' + str(seq_no)] = quantity

                seq_no += 1


 
        # pp.pprint(custom_order_list_dict)
        with open(file_name.replace('.xlsx', '_step_1.json'), "w") as json_file:
            json.dump(data_dict, json_file, indent = 4)




# append the detail info of each BOM item extract from Item List_step_1.json to generate BOM List_step_2.json file
# Expand the Customer order list item to all the BOM items involved into and filter the info based on the Interest_Duty_Class 
def step_2_processsing():

    bom_dict = {}
    for file_name in Dict_step_2.keys():
        with open(file_name) as json_file:
            Dict_step_2[file_name] = json.load(json_file)

    purchase_dict = {}

    for item_no, top_dict in Dict_step_2[WORKING_DIR + 'Purchase Lines_step_1.json'].items():
        tmp_dict = OrderedDict(sorted(top_dict['promise'].items()))
        purchase_dict[item_no] = {}
        purchase_dict[item_no]['promise'] = tmp_dict
        tmp_dict = OrderedDict(sorted(top_dict['request'].items()))
        purchase_dict[item_no]['request'] = tmp_dict

       
    with open(WORKING_DIR + 'Purchase Lines_step_2.json', "w") as json_file:
        json.dump(purchase_dict, json_file, indent = 4) 

    
    # test_list = []
    # for document_no, value in Dict_step_2['BOM list_step_1.json'].items():

    #     for key in value.keys():
    #         if key not in Dict_step_2['Item List_step_1.json']:
    #             print("BOM {} missing in item".format(key))

    #         test_list.append(key)


    #     tmp_dict = {}
    #     for bom_no, detail_dict in value.items():
    #         if bom_no not in Dict_step_2['Item List_step_1.json']:
    #             print("item_list err: no record for {}".format(bom_no))
    #         elif not Dict_step_2['Item List_step_1.json'][bom_no]['Duty Class']:
    #             print("item_list err: empty duty class for {}".format(bom_no))
    #         else:
    #             for key, value in Dict_step_2['Item List_step_1.json'][bom_no].items():
    #                 detail_dict.update({key : value})

    #             tmp_dict.update({bom_no:detail_dict})
        
    #     bom_dict.update({document_no:tmp_dict})

    # for key in Dict_step_2['Item List_step_1.json'].keys():
    #     if key not in test_list:
    #         print("item {} missing in BOM".format(key))    


    # with open('BOM list_step_2.json', "w") as json_file:
    #     json.dump(bom_dict, json_file, indent = 4)    

    bom_dict = Dict_step_2[WORKING_DIR + 'BOM list_step_1.json']
    filtered_dict = {}

    for key, value in Dict_step_2[WORKING_DIR + 'customer order list_step_1.json'].items():
        customer_document_no, bom_no = key.split('---')
        customer_dict = value
        # search the bom_no in the item list to determine if there is some storage for this to compensate the quantity
        if bom_no in Dict_step_2[WORKING_DIR + 'Item List_step_1.json']:
            if 'FG' in Dict_step_2[WORKING_DIR + 'Item List_step_1.json'][bom_no].get("Item Category Code", 'unknown'):
                on_hand_quantity = Dict_step_2[WORKING_DIR + 'Item List_step_1.json'][bom_no].get("Quantity on Hand", 0.0)
                Dict_step_2[WORKING_DIR + 'customer order list_step_1.json'][key].update({"Quantity on Hand from Item" : on_hand_quantity})                

        if bom_no in bom_dict:
            # print("there is associatekey in BOM for {}".format(value['Document No.']))
            bom_items_dict = bom_dict.get(bom_no, 'not found')

            for bom_item_no, detail_dict in bom_items_dict.items():

                value.update({bom_item_no:detail_dict})

            filtered_dict.update({key:value})
        else:
            print("customer order list err: {} no associate entry in BOM".format(bom_no))

    with open(WORKING_DIR + 'customer order list_step_2.json', "w") as json_file:
        json.dump(filtered_dict, json_file, indent = 4) 

# def step_2_1_processing():
#     with open('customer order list_step_2.json') as json_file:
#         custom_dict = json.load(json_file)

#     tmp_dict = {}
#     seq_no = 1
#     for combined_key, top_dict in custom_dict.items():
#         _, bom_no = combined_key.split('---')
#         new_combined_key = bom_no + '---' + top_dict["Shipment Date"] + '---' + str(seq_no)

#         tmp_dict[new_combined_key] = top_dict
#         seq_no += 1

#     combined_key_list = sorted(tmp_dict.keys())
#     bom_no_only_list = []
#     tmp_2_dict = {}
    
#     for combined_key in combined_key_list:
#         bom_no, _, _ = combined_key.split('---')
#         bom_no_only_list.append(bom_no)
    
#     bom_no_only_2_list = list(set(bom_no_only_list))


#     with open('customer order list_step_2_1.json', "w") as json_file:
#         json.dump(tmp_dict, json_file, indent = 4) 


    # dict_1 = OrderedDict(sorted(tmp_dict.items()))
    # bom_no_date_list = dict_1.keys()

    # for index in range(len(bom_no_date_list)):

    #     current_dict_key = bom_no_date_list[index]

    #     if float(dict_1[current_dict_key]["Quantity on Hand from Item"]) == 0.0 or 

    #     modified_record_list = []

    #     current_on_hand_quantity = float(dict_1[current_dict_key]["Quantity on Hand from Item"])
    #     # on hand quantity is not enough to cover the first customer order
    #     if current_on_hand_quantity <= float(dict_1[current_dict_key]["Quantity"]):
    #         dict_1[current_dict_key]["Quantity"] -= dict_1[current_dict_key]["Quantity on Hand from Item"]
    #         dict_1[current_dict_key]["Quantity on Hand from Item"] = 0

    #         # find all record with same bom_no with different date, and modify their on hand value to zero to indicate all 
    #         # on hand have been used by current order`
    #         bom_no, _ = current_dict_key.split('---')
    #         for offset in range(1, MAXIMUM_DUPLICATE_ORDER):
    #             next_record_dict_key = bom_no_date_list[index + offset]
    #             # same bom_no with variuos date
    #             if bom_no in next_record_dict_key:
    #                 dict_1[next_record_dict_key]["Quantity on Hand from Item"] = 0
    #             # end of iteration, because no more bom_no found
    #             else:
    #                 break
    #     # on hand quantity could cover the first customer order, make the quantity of current quantity to zero to indicate no demand
    #     # at all, but need cotinue this process to modify next customer order
    #     else:
            
    #         modified_record_list.append(index)
    #         dict_1[current_dict_key]["Quantity on Hand from Item"] -= dict_1[current_dict_key]["Quantity"] 
    #         dict_1[current_dict_key]["Quantity"] = 0

    #         bom_no, _ = current_dict_key.split('---')
    #         for offset in range(1, MAXIMUM_DUPLICATE_ORDER):
    #             next_record_dict_key = bom_no_date_list[index + offset]
    #             # same bom_no with variuos date
    #             if bom_no in next_record_dict_key:
    #                 dict_1[next_record_dict_key]["Quantity on Hand from Item"] = 0
    #             # end of iteration, because no more bom_no found
    #             else:
    #                 break


    # with open('customer order list_step_2_1.json', "w") as json_file:
    #     json.dump(dict_1, json_file, indent = 4) 

# use bom_no_id in the BOM List as the key to reorganize the customer order list to gather all customer order which belong to
# one bom item
def step_3_processing():

    with open(WORKING_DIR + 'customer order list_step_2.json') as json_file:
        custom_dict = json.load(json_file)

    final_dict = {}

    for custom_no, top_dict in custom_dict.items():

        # first iteration: search item dict 
        for item_no, level_2_dict in top_dict.items():
            if isinstance(level_2_dict, dict):
                tmp_dict = {}
                real_key = item_no

                # if materail not exist, create empty dict for adding
                if real_key not in final_dict:
                    final_dict.update({real_key:{}})

                # add bom detail info into dict
                for key, value in level_2_dict.items():
                    if "Description" in key :
                        final_dict[real_key].update({key:value})

                tmp_dict.update({"Production Quantity": level_2_dict["Quantity"]})
                tmp_dict.update({"Scrap %": level_2_dict["Scrap %"]})

                # after add bom info into the dict, should append customer order info
                final_dict[real_key].update({custom_no:tmp_dict})
                # append customer order info to each item entry
                for key, value in top_dict.items():
                    if not isinstance(value, dict):
                        final_dict[real_key][custom_no].update({key:value})                


    with open(WORKING_DIR + 'Item List_step_1.json') as json_file:
        item_dict = json.load(json_file)

    # append item info at top_dict for each item_no
    # find all items not showingup in BOM list but exist in the item list
    bom_item_list = []
    for bom_no, top_dict in final_dict.items():
        if bom_no in item_dict:
            for item_detail_key, value in item_dict[bom_no].items():
                top_dict[item_detail_key] = value
            bom_item_list.append(bom_no)

    for key, top_dict in item_dict.items():
        if key not in bom_item_list:
            final_dict[key] = top_dict

   
        # print("{} of duty '{}' showd up in {} customer orders".format(bom_key, top_dict['Duty Class'], order_cnt))
    
    with open(WORKING_DIR + 'customer order list_step_3.json', "w") as json_file:
        json.dump(final_dict, json_file, indent = 4)     

# print("now is {}".format(date_time.strftime(FMT)))



def write_to_xls_file():
    final_order_dict = OrderedDict()
    with open(WORKING_DIR + 'customer order list_step_3.json') as json_file:
        final_order_dict = json.load(json_file) 

    book = openpyxl.Workbook()
    sheet = book.active

    row_1_list = [  "Vendor",
                    'Production Code',
                    "Duty Class",
                    'Description',
                    "Base Unit of Measure",
                    "Total Demand",
                    # "Total on Hand",       
                    "Key Description",
                    "Past Due",
                ]

    specific_column_list = [
                        'On Hand (W1)',
                        'On Hand (W2)',
                        'Vendor Floor Stock',
                        'All Open PO & BPO',
                        'Purchases Receipts',
                        'Demand',
                        'Available Inventory',
                        'Total Ending Balance',   
    ]

    future_4_weeks_day_list = []

    header_font = Font(color="000000")
    sub_row_font = [
                        Font(color="000000"),
                        Font(color="000000"),
                        Font(color="000000"),
                        Font(color="000000"),
                        Font(color="000000"),
                        Font(color="000000"),
                        Font(color="000000"),
                        Font(color="000000"),
                    ]


    for offset in range(14):
        next_day = Current_day + timedelta(days=offset+1)
        row_1_list.append(next_day.strftime(FMT))

    for offset_days in range(1, 29):
        if (offset_days % 7) == 0:
            next_week_day = next_day + timedelta(days=offset_days)
            row_1_list.append(next_week_day.strftime(FMT))  
        future_4_weeks_day_list.append((next_day + timedelta(days=offset_days)).strftime(FMT))  

    print("row list is {}".format(row_1_list))

    row_cnt = 1
    column_cnt = 1

    for column_cnt in range(len(row_1_list)):
        sheet.cell(row=row_cnt, column=column_cnt+1).value = row_1_list[column_cnt]
        sheet.cell(row=row_cnt, column=column_cnt+1).font = header_font
        try:
            date = datetime.strptime(row_1_list[column_cnt], FMT)
            year =date.strftime(YEAR_FMT)

            if date.date() in holidays.Canada(years = int(year)).keys() or date.weekday() > 4:
                sheet.cell(row=row_cnt, column=column_cnt+1).fill = PatternFill(start_color="B3F2FF", end_color="B3F2FF", fill_type = "solid")
        except:
            sheet.cell(row=row_cnt, column=column_cnt+1).fill = PatternFill(start_color="B3F2FF", end_color="B3F2FF", fill_type = "solid")

    sheet.freeze_panes = "A2"

    sheet.print_title_rows='1:1'

    row_cnt += 1
    column_cnt = 1

    with open(WORKING_DIR + 'Purchase Lines_step_2.json') as json_file:
        purchase_dict = json.load(json_file)

    for bom_key, top_dict in final_order_dict.items():
        ttt_dict = {}
        if top_dict.get("Duty Class", "unknown") in Interest_Duty_List:

            demand_dict = defaultdict(lambda: float(0))
            demand_dict_tmp = defaultdict(lambda: float(0))

            if bom_key in purchase_dict:
                open_po_tmp_dict = {}
                promised_date_list = []
                request_date_list = []
                purchase_promise_dict = purchase_dict[bom_key]['promise']
                for date_seq, quantity in purchase_promise_dict.items():
                    date, _ = date_seq.split('---')
                    
                    if date not in promised_date_list:
                        promised_date_list.append(date)
                        
                        open_po_tmp_dict[date] = quantity
                    else:
                        open_po_tmp_dict[date] += quantity

                no_week_day_list = []
                for day in row_1_list[0:-4]:
                    no_week_day_list.append(day)

                open_po_dict = defaultdict(lambda: float(0))
                # extra processing to accumulate the demand in the future 4 weeks
                for timestamp in open_po_tmp_dict.keys():
                    if timestamp in future_4_weeks_day_list:
                        time_diff = datetime.strptime(timestamp, FMT) - datetime.strptime(future_4_weeks_day_list[0], FMT)
                        if timedelta(days=0) < time_diff <= timedelta(days=6):
                            open_po_dict[future_4_weeks_day_list[6]] +=open_po_tmp_dict[timestamp]

                        elif timedelta(days=6) < time_diff <= timedelta(days=13):
                            open_po_dict[future_4_weeks_day_list[13]] +=open_po_tmp_dict[timestamp]
                        elif timedelta(days=13) < time_diff <= timedelta(days=20):
                            open_po_dict[future_4_weeks_day_list[20]] +=open_po_tmp_dict[timestamp]
                        elif timedelta(days=20) < time_diff <= timedelta(days=27):
                            open_po_dict[future_4_weeks_day_list[27]] +=open_po_tmp_dict[timestamp]
                    elif timestamp not in no_week_day_list:
                        open_po_dict['Past Due'] += open_po_tmp_dict[timestamp]
                    else:
                        open_po_dict[timestamp] = open_po_tmp_dict[timestamp]

                purchase_receipt_tmp_dict = {}
                request_date_list = []
                purchase_request_dict = purchase_dict[bom_key]['request']
                for date_seq, quantity in purchase_request_dict.items():
                    date, _ = date_seq.split('---')
                    
                    if date not in request_date_list:
                        request_date_list.append(date)
                        
                        purchase_receipt_tmp_dict[date] = quantity
                    else:
                        purchase_receipt_tmp_dict[date] += quantity

                purchase_receipt_dict = defaultdict(lambda: float(0))

                # extra processing to accumulate the demand in the future 4 weeks
                for timestamp in purchase_receipt_tmp_dict.keys():
                    if timestamp in future_4_weeks_day_list:
                        time_diff = datetime.strptime(timestamp, FMT) - datetime.strptime(future_4_weeks_day_list[0], FMT)
                        if timedelta(days=0) < time_diff <= timedelta(days=6):
                            purchase_receipt_dict[future_4_weeks_day_list[6]] +=purchase_receipt_tmp_dict[timestamp]

                        elif timedelta(days=6) < time_diff <= timedelta(days=13):
                            purchase_receipt_dict[future_4_weeks_day_list[13]] +=purchase_receipt_tmp_dict[timestamp]
                        elif timedelta(days=13) < time_diff <= timedelta(days=20):
                            purchase_receipt_dict[future_4_weeks_day_list[20]] +=purchase_receipt_tmp_dict[timestamp]
                        elif timedelta(days=20) < time_diff <= timedelta(days=27):
                            purchase_receipt_dict[future_4_weeks_day_list[27]] +=purchase_receipt_tmp_dict[timestamp]
                    elif timestamp not in no_week_day_list:
                        purchase_receipt_dict['Past Due'] += purchase_receipt_tmp_dict[timestamp]
                    else:
                        purchase_receipt_dict[timestamp] = purchase_receipt_tmp_dict[timestamp]


            for middle_key, middle_dict in top_dict.items():
                if isinstance(middle_dict, dict):
                    per_unit_production_quantity = float(middle_dict["Production Quantity"])
                    scrap = float(middle_dict["Scrap %"])
                    order_quantity = float(middle_dict["Outstanding Quantity"])
                    demand_quantity = order_quantity * per_unit_production_quantity * (1.0 + scrap * 0.01)

                    # print("demand for bom:{} on order {} '{}'   is   {:.3f}        = {}    *   {}*    (1.0 + {} * 0.01)".format(bom_key, middle_key, middle_dict["Shipment Date"], demand_quantity, 
                    #                         order_quantity, 
                    #                         per_unit_production_quantity, 
                    #                         scrap))

                    # ttt_dict[middle_dict["Shipment Date"] + '-' +  middle_key] = "{:.3f}  = {} * {} * (1.0 + {} * 0.01)".format(demand_quantity, order_quantity, per_unit_production_quantity,  scrap)
                    
                    # till now only occur once
                    if middle_dict["Shipment Date"] not in demand_dict_tmp:
                        demand_dict_tmp.update({middle_dict["Shipment Date"]: demand_quantity})
                    # already have one, accumulate the demand on it
                    else:
                        demand_sum = demand_dict_tmp[middle_dict["Shipment Date"]] + demand_quantity
                        demand_dict_tmp.update({middle_dict["Shipment Date"]: demand_sum})

            # print("after analyze same date for shipment this is the result ")
            # pprint.pprint(dict(demand_dict_tmp))
            # pprint.pprint(dict(sorted(ttt_dict.items())))


            # print('tmp_dict is {}'.format(demand_dict_tmp))

            # extra processing to accumulate the demand in the future 4 weeks
            for timestamp in demand_dict_tmp.keys():
                if timestamp in future_4_weeks_day_list:
                    time_diff = datetime.strptime(timestamp, FMT) - datetime.strptime(future_4_weeks_day_list[0], FMT)
                    if timedelta(days=0) < time_diff <= timedelta(days=6):
                        demand_dict[future_4_weeks_day_list[6]] +=demand_dict_tmp[timestamp]

                    elif timedelta(days=6) < time_diff <= timedelta(days=13):
                        demand_dict[future_4_weeks_day_list[13]] +=demand_dict_tmp[timestamp]
                    elif timedelta(days=13) < time_diff <= timedelta(days=20):
                        demand_dict[future_4_weeks_day_list[20]] +=demand_dict_tmp[timestamp]
                    elif timedelta(days=20) < time_diff <= timedelta(days=27):
                        demand_dict[future_4_weeks_day_list[27]] +=demand_dict_tmp[timestamp]

                
            # print('dict is {}'.format(demand_dict))
            # add record of each signle day into the lookup dictionary
            for key, value in demand_dict_tmp.items():
                demand_dict.update({key: value})



            # here put each value into cell
            # loop for 8 sub_row for each item no.                        
            for sub_row in range(8):
                sheet.cell(row=row_cnt, column=1).value = top_dict.get("Vendor No.", "unknown")
                sheet.cell(row=row_cnt, column=2).value = bom_key
                sheet.cell(row=row_cnt, column=3).value = top_dict.get("Duty Class", "uknown")
                sheet.cell(row=row_cnt, column=4).value = top_dict.get('Description', "uknown")
                sheet.cell(row=row_cnt, column=5).value = top_dict.get("Base Unit of Measure", "uknown")

                sheet.cell(row=row_cnt, column=7).value = specific_column_list[sub_row]
                sheet.cell(row=row_cnt, column=7).font = sub_row_font[sub_row]

                total_demand = 0

                # add demand number into the excel
                if 'Demand' in specific_column_list[sub_row]:
                    # put demand value for each following 14 days
                    for column_no in range(7, len(row_1_list)):
                        if row_1_list[column_no] in demand_dict:
                            sheet.cell(row=row_cnt, column=column_no + 1).value = int(round(demand_dict[row_1_list[column_no]]))
                            total_demand += sheet.cell(row=row_cnt, column=column_no + 1).value

                elif 'All Open PO & BPO'in specific_column_list[sub_row]:
                    # put demand value for each following 14 days
                    for column_no in range(7, len(row_1_list)):
                        if row_1_list[column_no] in open_po_dict:
                            sheet.cell(row=row_cnt, column=column_no + 1).value = int(round(purchase_receipt_dict[row_1_list[column_no]]))

                elif 'Purchases Receipts'in specific_column_list[sub_row]:
                    # put demand value for each following 14 days
                    for column_no in range(7, len(row_1_list)):
                        if row_1_list[column_no] in purchase_receipt_dict:
                            sheet.cell(row=row_cnt, column=column_no + 1).value = int(round(open_po_dict[row_1_list[column_no]]))

                
                elif 'On Hand (W1)' in specific_column_list[sub_row]:
                    sheet.cell(row=row_cnt, column=8).value = int(top_dict.get("W1", 0))

                elif 'On Hand (W2)' in specific_column_list[sub_row]:
                    sheet.cell(row=row_cnt, column=8).value = int(top_dict.get("W2", 0))

                sheet.cell(row=row_cnt, column=6).value = int(round(total_demand))

                row_cnt += 1
                column_cnt = 1
        
        column_cnt = 1

    print("write {} row".format(row_cnt))

                
    book.save(WORKING_DIR + "output_{}.xlsx".format(datetime.now().strftime(FMT)))


# # read_data_from_files(file_dict)
# step_2_processsing()
# # step_2_1_processing()
# step_3_processing()
# write_to_xls_file()

def build_parser():
    '''
    Build command line options.
    '''
    parser = argparse.ArgumentParser(
        description='Procesing the xls files')

    # Positional Arguments
    parser.add_argument("-d", "--dst_dir",
                        type=str,
                        help='destination direcoty')

    parser.add_argument("-r", "--running_steps",
                        type=str,
                        help='all steps or skip first reading which consume most of time <all|skip>')

    return parser

def main():
    '''
    Args:
        argv[1](str) :  destination directory
        argv[2](str) :  step all or exclude step_1 of reading inital execl file

    Returns:
    '''
    

    # Load argparse and parse arguments.
    cmd_parser = build_parser()
    args = cmd_parser.parse_args(sys.argv[1:])


    #Commnd from Zabbix server to query the RPD stats
    WORKING_DIR = args.dst_dir

    if 'skip' in args.running_steps:
    # read_data_from_files(file_dict)
        step_2_processsing()
        # step_2_1_processing()
        step_3_processing()
        write_to_xls_file()
    elif 'all' in args.running_steps:
        read_data_from_files(file_dict)
        step_2_processsing()
        # step_2_1_processing()
        step_3_processing()
        write_to_xls_file()

if __name__ == '__main__':
    main()
