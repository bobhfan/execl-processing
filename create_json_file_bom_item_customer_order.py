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

# time = datetime.strptime("2021-02-13 00:00:00", FMT)

# date_time = datetime.now()

# print("now is {}".format(date_time.strftime(FMT)))

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
                        if 'RA 266' in item_no:
                            print("got it")
                            import pdb; pdb.set_trace()

                    elif 'Document No.' in field_name_dict[colum_n] :
                        document_no = sheet.cell(row=line, column=colum_n).value

                    elif "Promised Receipt Date" in field_name_dict[colum_n] :
                        if sheet.cell(row=line, column=colum_n).value:
                            promise_date = sheet.cell(row=line, column=colum_n).value.strftime(FMT)
                        else:
                            promise_date = ''

                    elif "Outstanding Quantity" in field_name_dict[colum_n] :
                        quantity = sheet.cell(row=line, column=colum_n).value

                    elif "Requested Receipt Date" in field_name_dict[colum_n] :
                        if sheet.cell(row=line, column=colum_n).value and isinstance(sheet.cell(row=line, column=colum_n).value, datetime):
                            request_date = sheet.cell(row=line, column=colum_n).value.strftime(FMT)
                        else:
                            request_date = ''

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

                if 'promise_date' in locals() and promise_date:
                    data_dict[item_no]['promise'][promise_date + '---' + str(seq_no)] = quantity
                elif 'request_date' in locals() and request_date:
                    data_dict[item_no]['request'][request_date + '---' + str(seq_no)] = quantity

                seq_no += 1

        # pp.pprint(custom_order_list_dict)
        with open(file_name.replace('.xlsx', '_step_1.json'), "w") as json_file:
            json.dump(data_dict, json_file, indent = 4)




# append the detail info of each BOM item extract from Item List_step_1.json to generate BOM List_step_2.json file
# Expand the Customer order list item to all the BOM items involved into and filter the info based on the Interest_Duty_Class
def step_2_processsing(dict_step_2):

    bom_dict = {}
    for file_name in dict_step_2.keys():
        if "BOM" in file_name:
            bom_name = file_name
        elif "Item" in file_name:
            item_name = file_name
        elif "Purchase" in file_name:
            purchase_name = file_name
        elif "customer" in file_name:
            customer_name = file_name
        with open(file_name) as json_file:
            dict_step_2[file_name] = json.load(json_file)

    purchase_dict = {}

    for item_no, top_dict in dict_step_2[purchase_name].items():
        tmp_dict = OrderedDict(sorted(top_dict['promise'].items()))
        purchase_dict[item_no] = {}
        purchase_dict[item_no]['promise'] = tmp_dict
        tmp_dict = OrderedDict(sorted(top_dict['request'].items()))
        purchase_dict[item_no]['request'] = tmp_dict


    with open(purchase_name.replace('_1.j', '_2.j'), "w") as json_file:
        json.dump(purchase_dict, json_file, indent = 4)

    bom_dict = dict_step_2[bom_name]
    filtered_dict = {}

    for key, value in dict_step_2[customer_name].items():
        customer_document_no, bom_no = key.split('---')
        customer_dict = value
        # search the bom_no in the item list to determine if there is some storage for this to compensate the quantity
        if bom_no in dict_step_2[item_name]:
            if 'FG' in dict_step_2[item_name][bom_no].get("Item Category Code", 'unknown'):
                on_hand_quantity = dict_step_2[item_name][bom_no].get("Quantity on Hand", 0.0)
                dict_step_2[customer_name][key].update({"Quantity on Hand from Item" : on_hand_quantity})

        if bom_no in bom_dict:
            # print("there is associatekey in BOM for {}".format(value['Document No.']))
            bom_items_dict = bom_dict.get(bom_no, 'not found')

            for bom_item_no, detail_dict in bom_items_dict.items():

                value.update({bom_item_no:detail_dict})

            filtered_dict.update({key:value})
        else:
            print("customer order list err: {} no associate entry in BOM".format(bom_no))

    with open(customer_name.replace('_1.j', '_2.j'), "w") as json_file:
        json.dump(filtered_dict, json_file, indent = 4)

def step_2_1_processing():
    with open('./0527/customer order list_step_2.json') as json_file:
        custom_dict = json.load(json_file)

    tmp_dict = {}
    seq_no = 1

    bom_no_list =[combined_key.split('---')[-1] for combined_key in custom_dict.keys()]

    print("all customer order is {}".format(len(bom_no_list)))

    bom_no_no_duplicate_list = list(set(bom_no_list))

    print("bom number involved is {}".format(len(bom_no_no_duplicate_list)))

    # find all order with same bom_no and sorted by ship date
    for key in bom_no_no_duplicate_list:
        tmp_dict[key] = []
        for dict_key, top_dict in custom_dict.items():
            ship_date = top_dict["Shipment Date"]
            if key == dict_key.split('---')[-1]:
                customer_no, bom_no = dict_key.split('---')
                tmp_dict[key].append('---'.join((bom_no, ship_date, customer_no)))
        
        tmp_dict[key].sort()

    tmp_2_dict = {}
    for bom_no, combinde_key_list in tmp_dict.items():
        tmp_2_dict[bom_no] = []
        for combined_key in combinde_key_list:
            bom_no_1, ship_date, customer_no = combined_key.split('---')
            # if bom_no == bom_no_1:
            search_key = '---'.join((customer_no, bom_no))
            order_quantiry = custom_dict[search_key]['Quantity']
            quantity_on_hand = custom_dict[search_key].get('Quantity on Hand from Item', 0)
            tmp_2_dict[bom_no].append((combined_key, order_quantiry, quantity_on_hand))

    for bom_no, entries in tmp_2_dict.items():
        sum_of_quantity = 0
        for index in range(len(entries)):
            sum_of_quantity += entries[index][1]
            if sum_of_quantity >= entries[0][-1]:
                rest_quantity = sum_of_quantity - entries[0][-1]
                break
        
        if index < len(entries):
            print("bom_no: {} only [0 : {}] need be changed and the rest is {}".format(bom_no, index, rest_quantity))
        else:
            print("bom_no: {} all of element need be changed".format(bom_no))



    with open('./0527/customer order list_step_2_1.json', "w") as json_file:
        json.dump(tmp_2_dict, json_file, indent = 4)


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


# use bom_no_id in the BOM List as the key to reorganize the customer order list to gather all customer order which belong to
# one bom item
def step_3_processing(file_name_step_3_dict):

    with open(file_name_step_3_dict['customer_name']) as json_file:
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


    with open(file_name_step_3_dict['item_name']) as json_file:
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

    with open(file_name_step_3_dict['customer_name'].replace('_2.j', '_3.j'), "w") as json_file:
        json.dump(final_dict, json_file, indent = 4)

# print("now is {}".format(date_time.strftime(FMT)))


def generate_puchase_data(input_dict):
    exist_date_list = []
    output_dict = {}
    for date_seq, quantity in input_dict.items():
        date, _ = date_seq.split('---')

        if date not in exist_date_list:
            exist_date_list.append(date)

            output_dict[date] = quantity
        else:
            output_dict[date] += quantity

    return output_dict

def distribute_data_to_different_period(input_dict, days_list, future_4_weeks_day_list):

    single_days_list = days_list[0:-4]
    output_dict = {}

    output_dict = defaultdict(lambda: float(0))
    # extra processing to accumulate the demand in the future 4 weeks
    for timestamp in input_dict.keys():
        if timestamp in future_4_weeks_day_list:
            time_diff = datetime.strptime(timestamp, FMT).date() - datetime.strptime(future_4_weeks_day_list[0], FMT).date()
            if timedelta(days=0) <= time_diff <= timedelta(days=6):
                output_dict[future_4_weeks_day_list[6]] += input_dict[timestamp]

            elif timedelta(days=6) < time_diff <= timedelta(days=13):
                output_dict[future_4_weeks_day_list[13]] += input_dict[timestamp]
            elif timedelta(days=13) < time_diff <= timedelta(days=20):
                output_dict[future_4_weeks_day_list[20]] += input_dict[timestamp]
            elif timedelta(days=20) < time_diff <= timedelta(days=27):
                output_dict[future_4_weeks_day_list[27]] += input_dict[timestamp]
        elif timestamp not in single_days_list:
            output_dict['Past Due'] += input_dict[timestamp]
        else:
            output_dict[timestamp] = input_dict[timestamp]

    return output_dict



def write_to_xls_file(file_name_step_4_dict):
    final_order_dict = OrderedDict()
    with open(file_name_step_4_dict['customer_name']) as json_file:
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

        if column_cnt > len(row_1_list) - 5 :
            sheet.cell(row=row_cnt, column=column_cnt+1).fill = PatternFill(start_color="CDD1D6", end_color="CDD1D6", fill_type = "solid")
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

    with open(file_name_step_4_dict['purchase_name']) as json_file:
        purchase_dict = json.load(json_file)

    for bom_key, top_dict in final_order_dict.items():
        ttt_dict = {}
        if bom_key == "RA 263" or bom_key == "RA 254":
            print("Got RA 263 or 254")

        if top_dict.get("Duty Class", "unknown") in Interest_Duty_List:

            demand_dict = defaultdict(lambda: float(0))
            demand_dict_tmp = defaultdict(lambda: float(0))


            purchase_promise_dict = {}
            purchase_request_dict = {}

            if bom_key in purchase_dict:

                purchase_promise_intermediate_dict = generate_puchase_data(purchase_dict[bom_key]['promise'])
                purchase_promise_dict = distribute_data_to_different_period(purchase_promise_intermediate_dict, row_1_list, future_4_weeks_day_list)

                purchase_request_intermediate_dict = generate_puchase_data(purchase_dict[bom_key]['request'])
                purchase_request_dict = distribute_data_to_different_period(purchase_request_intermediate_dict, row_1_list, future_4_weeks_day_list)


            for middle_key, middle_dict in top_dict.items():
                if isinstance(middle_dict, dict):
                    per_unit_production_quantity = float(middle_dict["Production Quantity"])
                    scrap = float(middle_dict["Scrap %"])
                    order_quantity = float(middle_dict["Outstanding Quantity"])
                    demand_quantity = order_quantity * per_unit_production_quantity * (1.0 + scrap * 0.01)

                    # till now only occur once
                    if middle_dict["Shipment Date"] not in demand_dict_tmp:
                        demand_dict_tmp.update({middle_dict["Shipment Date"]: demand_quantity})
                    # already have one, accumulate the demand on it
                    else:
                        demand_sum = demand_dict_tmp[middle_dict["Shipment Date"]] + demand_quantity
                        demand_dict_tmp.update({middle_dict["Shipment Date"]: demand_sum})

            demand_dict = distribute_data_to_different_period(demand_dict_tmp, row_1_list, future_4_weeks_day_list)

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
                        if row_1_list[column_no] in purchase_request_dict:
                            sheet.cell(row=row_cnt, column=column_no + 1).value = int(round(purchase_request_dict[row_1_list[column_no]]))

                elif 'Purchases Receipts'in specific_column_list[sub_row]:
                    # put demand value for each following 14 days
                    for column_no in range(7, len(row_1_list)):
                        if row_1_list[column_no] in purchase_promise_dict:
                            sheet.cell(row=row_cnt, column=column_no + 1).value = int(round(purchase_promise_dict[row_1_list[column_no]]))
                            sheet.cell(row=row_cnt, column=column_no + 1).font = Font( color='00004D', bold=True, italic=True)

                elif 'On Hand (W1)' in specific_column_list[sub_row]:
                    sheet.cell(row=row_cnt, column=8).value = int(top_dict.get("W1", 0))

                elif 'On Hand (W2)' in specific_column_list[sub_row]:
                    sheet.cell(row=row_cnt, column=8).value = int(top_dict.get("W2", 0))

                sheet.cell(row=row_cnt, column=6).value = int(round(total_demand))

                row_cnt += 1
                column_cnt = 1

            # after all value could be read from the dict directly, and prepared moving to nenext iteration for next bmo_no, we could start calculating
            # some cell based on the cell value
            day_column_start = 9
            past_due_column = 8
            previous_w1_past_due_row = row_cnt - 8


            previous_purchase_receipts_row = previous_w1_past_due_row + 4
            previous_demand_row = previous_w1_past_due_row + 5
            previous_available_inventory_row = previous_w1_past_due_row + 6
            previous_total_ending_balance_row = previous_w1_past_due_row + 7
            
            # cell(row=Available Inventory, column=Past due) ---> W1 + purchasde Receipts 
            if not sheet.cell(row=previous_w1_past_due_row, column=past_due_column).value:
                sheet.cell(row=previous_w1_past_due_row, column=past_due_column).value = 0
            if not sheet.cell(row=previous_purchase_receipts_row, column=past_due_column).value:
                sheet.cell(row=previous_purchase_receipts_row, column=past_due_column).value = 0
            if not sheet.cell(row=previous_demand_row, column=past_due_column).value:
                sheet.cell(row=previous_demand_row, column=past_due_column).value = 0


            sheet.cell(row=previous_available_inventory_row, column=past_due_column).value = \
                int(sheet.cell(row=previous_w1_past_due_row, column=past_due_column).value) - \
                int(sheet.cell(row=previous_demand_row, column=past_due_column).value)

            # cell(row=Total Ending balance,column=past Due) ---> W1 + purchasde Receipts - demand
            sheet.cell(row=previous_total_ending_balance_row, column=past_due_column).value = \
                int(sheet.cell(row=previous_available_inventory_row, column=past_due_column).value) + \
                int(sheet.cell(row=previous_purchase_receipts_row, column=past_due_column).value)            

            for last_2_row_column_no in range(9, len(row_1_list)):
                if not sheet.cell(row=previous_demand_row, column=last_2_row_column_no).value:
                    sheet.cell(row=previous_demand_row, column=last_2_row_column_no).value = 0
                if not sheet.cell(row=previous_purchase_receipts_row, column=last_2_row_column_no).value:
                    sheet.cell(row=previous_purchase_receipts_row, column=last_2_row_column_no).value = 0

                sheet.cell(row=previous_available_inventory_row, column=last_2_row_column_no).value = \
                    int(sheet.cell(row=previous_available_inventory_row, column=last_2_row_column_no - 1).value) - \
                    int(sheet.cell(row=previous_demand_row, column=last_2_row_column_no).value)

                sheet.cell(row=previous_total_ending_balance_row, column=last_2_row_column_no).value = \
                    int(sheet.cell(row=previous_available_inventory_row, column=last_2_row_column_no).value) + \
                    int(sheet.cell(row=previous_purchase_receipts_row, column=last_2_row_column_no).value)

        column_cnt = 1

    print("write {} row".format(row_cnt))


    book.save("output_{}.xlsx".format(datetime.now().strftime(FMT)))


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
    working_dir = args.dst_dir

    file_dict = {
                # working_dir + 'BOM list.xlsx' : 1,
                # working_dir + 'Item List.xlsx' : 1,
                # working_dir + 'customer order list.xlsx': 1,
                working_dir + 'Purchase Lines.xlsx':1,
            }

    dict_step_2 = {
                working_dir + 'BOM list_step_1.json' : {},
                working_dir + 'Item List_step_1.json' : {},
                working_dir + 'customer order list_step_1.json': {},
                working_dir + 'Purchase Lines_step_1.json' : {},
                }

    file_name_step_3_dict = {
                            'customer_name' : working_dir + 'customer order list_step_2.json',
                            'item_name' : working_dir + 'Item List_step_1.json',
    }

    file_name_step_4_dict = {
                            'customer_name' : working_dir + 'customer order list_step_3.json',
                            'purchase_name' : working_dir + 'Purchase Lines_step_2.json',
    }

    if 'skip' in args.running_steps:
        step_2_processsing(dict_step_2)
        # step_2_1_processing()
        step_3_processing(file_name_step_3_dict)
        write_to_xls_file(file_name_step_4_dict)

    elif 'all' in args.running_steps:
        read_data_from_files(file_dict)
        step_2_processsing(dict_step_2)
        # step_2_1_processing()
        step_3_processing(file_name_step_3_dict)
        write_to_xls_file(file_name_step_4_dict)

if __name__ == '__main__':
    main()
