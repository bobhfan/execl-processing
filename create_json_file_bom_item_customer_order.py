import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta
import pprint
import json
from collections import OrderedDict, defaultdict

FMT = '%Y-%m-%d'


# time = datetime.strptime("2021-02-13 00:00:00", FMT)

# date_time = datetime.now()

# print("now is {}".format(date_time.strftime(FMT)))
file_dict = {
                'BOM list.xlsx' : 1,
                'Item List.xlsx' : 1,
                'customer order list.xlsx': 1,
                'Purchase Lines.xlsx':1,
            }


Item_List_Interest_Field_List = ['Duty Class', 
                                'W1',
                                'W2',
                                'Base Unit of Measure' ,
                                'Vendor No.',
                                ]

Interest_Duty_List = ["FILM", "BAG",  "SEASONING"]

Valid_Purchase_Period = timedelta(days=365)

Purchase_Lines_Interest_Dict = {
        'Promised Receipt Date':'Promised Receipt Date',
        'Quantity' : 'Purchase Quantity',
        'Requested Receipt Date' : 'Requested Receipt Date',
        'Outstanding Quantity' : 'Outstanding Quantity' ,
    }

Dict_step_2 = {
                'BOM list_step_1.json' : {},
                'Item List_step_1.json' : {},
                'customer order list_step_1.json': {},
                'Purchase Lines_step_1.json' : {},
                }

def read_data_from_files(file_dict):

    for file_name, start_row in file_dict.items():
        wb= openpyxl.load_workbook(file_name)
        names = wb.sheetnames
        print(wb.sheetnames)
        sheet = wb.active

        data_dict = {}
        field_name_dict = {}

        for x in range (start_row,start_row + 1):

            for y in range(1,sheet.max_column + 1):
                field_name_dict.update({y:sheet.cell(row=x,column=y).value})
            # data_dict.update({'row_1':field_name_dict})
        pp = pprint.PrettyPrinter(indent=4)
        pp.pprint(field_name_dict)


        # # iterate through excel and display data
        print("{}'s total row is {}".format(file_name, sheet.max_row))

        for line in range(start_row + 1, sheet.max_row+1):
        # for line in range(3, 155):
            tmp_dict = {}
            for colum_n in range(1, sheet.max_column+1):
                # This is for generate the dict which key is the value of 'No.' and value is the  
                # value of "Duty Class" to help spped up the lookup of the material covered by BOM list

                if 'Item List.xlsx' in file_name :
                    if 'No.' in field_name_dict[colum_n] \
                        and ' ' not in field_name_dict[colum_n]:
                        
                        item_no = sheet.cell(row=line, column=colum_n).value

                    elif field_name_dict[colum_n] in Item_List_Interest_Field_List:
                        tmp_dict.update({field_name_dict[colum_n] : sheet.cell(row=line, column=colum_n).value}) 

                    elif 'Blocked' in field_name_dict[colum_n]:
                        item_block = sheet.cell(row=line, column=colum_n).value

                if 'Purchase Lines.xlsx' in file_name :
                    if 'No.' in field_name_dict[colum_n] and ' ' not in field_name_dict[colum_n] :
                        item_no = sheet.cell(row=line, column=colum_n).value

                    elif 'Document No.' in field_name_dict[colum_n] :
                        document_no = sheet.cell(row=line, column=colum_n).value

                    elif field_name_dict[colum_n] in Purchase_Lines_Interest_Dict.keys():
                        if isinstance(sheet.cell(row=line, column=colum_n).value, datetime):
                            tmp_dict.update(
                                {field_name_dict[colum_n] : 
                                sheet.cell(row=line, column=colum_n).value.strftime(FMT)})
                        else:
                            tmp_dict.update({Purchase_Lines_Interest_Dict[field_name_dict[colum_n]] : sheet.cell(row=line, column=colum_n).value}) 


                # This is for generate the dict which key is the value of 'Production BOM No.' and the dict is
                # other field
                elif 'BOM list.xlsx' in file_name :

                    
                    if 'Production BOM No.' in field_name_dict[colum_n]:
                        dict_key = sheet.cell(row=line, column=colum_n).value 
                        # have this code first time should create empty dict for further
  

                    elif 'No.' in field_name_dict[colum_n] and ' ' not in field_name_dict[colum_n]:
                        material_dict_key = sheet.cell(row=line, column=colum_n).value 

                    else:
                        if isinstance(sheet.cell(row=line, column=colum_n).value, datetime):
                            tmp_dict.update(
                                {field_name_dict[colum_n] : 
                                sheet.cell(row=line, column=colum_n).value.strftime(FMT)})
                        else:

                            tmp_dict.update({
                                field_name_dict[colum_n] : 
                                sheet.cell(row=line, column=colum_n).value})     
                
                # This is for generate the dict which key is the value of 'No.' and the dict is
                # other fields
                elif 'customer order list.xlsx' in file_name :
                    if 'No.' in field_name_dict[colum_n] \
                        and ' ' not in field_name_dict[colum_n]:
                        dict_key = sheet.cell(row=line, column=colum_n).value     

                    elif 'Document No.' in field_name_dict[colum_n]:
                        customer_document_no = sheet.cell(row=line, column=colum_n).value  

                    elif 'Quantity' in field_name_dict[colum_n] and not ' ' in field_name_dict[colum_n]:
                        tmp_dict.update({
                                'Customer order Quantity' : 
                                sheet.cell(row=line, column=colum_n).value})      

                    else:
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
                if 'No' in item_block:
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
                    data_dict.update({item_no:{}})
                
                data_dict[item_no].update({
                    document_no : 
                    tmp_dict})


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

    for item_no, top_dict in Dict_step_2['Purchase Lines_step_1.json'].items():
        tmp_promise_dict = {}
        tmp_expect_dict = {}
        tmp_quantiry_dict = {}
        tmp_dict = {}
        for po_no, second_dict in top_dict.items():
            for key, value in second_dict.items():
                if "Promised" in key:
                    if not value:
                        value = '1976-01-01'
                    tmp_promise_dict.update({po_no + key : value})
                elif "Expected" in key:
                    if not value:
                        value = '1976-01-01'
                    tmp_expect_dict.update({po_no + key : value})
                else:
                    tmp_quantiry_dict.update({po_no + key : value})
        marklist=sorted((value, key) for (key,value) in tmp_promise_dict.items()
            if datetime.now() - datetime.strptime(value, FMT) < Valid_Purchase_Period)
        sortdict=dict([(k,v) for v,k in marklist])
        purchase_dict.update({item_no:{}})
        purchase_dict[item_no]['promise'] = sortdict


        marklist=sorted((value, key) for (key,value) in tmp_expect_dict.items()
            if datetime.now() - datetime.strptime(value, FMT) < Valid_Purchase_Period)
        sortdict=dict([(k,v) for v,k in marklist])
        purchase_dict[item_no]['expect'] = sortdict

        purchase_dict[item_no]['quantity'] = tmp_quantiry_dict
        
    with open('Purchase Lines_step_22.json', "w") as json_file:
        json.dump(purchase_dict, json_file, indent = 4) 
        
        po_no_list = []
        po_promise_date_list =[]
        po_expect_date_list = []

        for po_no, second_dict in top_dict.items():
            po_no_list.append(po_no)
            if not second_dict.get("Promised Receipt Date"):
                po_promise_date_list.append('1976-01-01')
            else :
                po_promise_date_list.append(second_dict.get("Promised Receipt Date"))
            if not second_dict.get("Expected Receipt Date"):
                po_expect_date_list.append('1976-01-01')
            else :
                po_expect_date_list.append(second_dict.get("Expected Receipt Date"))


        latest_po_promise_date = max(po_promise_date_list)
        latest_po_expect_date = max(po_expect_date_list)

        promise_latest_index_list = []
        expect_latest_index_list = []



        if datetime.now() - datetime.strptime(latest_po_promise_date, FMT) < Valid_Purchase_Period:
            for index in range(len(po_promise_date_list)):
                if latest_po_promise_date == po_promise_date_list[index] and latest_po_promise_date != '0-0-0':
                    promise_latest_index_list.append(index)
            if len(promise_latest_index_list) > 1:
                print("the promise list is {}".format(po_promise_date_list))
                print("process purchase data error, there are two identical latest promise receipt date for {}".format(
                        item_no
                ))
        
        if datetime.now() - datetime.strptime(latest_po_expect_date, FMT)  < Valid_Purchase_Period:
            for index in range(len(po_expect_date_list)):
                if latest_po_expect_date == po_expect_date_list[index] and latest_po_expect_date != '0-0-0' :
                    expect_latest_index_list.append(index)
            if len(expect_latest_index_list) > 1:
                print("the expect list is {}".format(po_expect_date_list))
                print("process purchase data error, there are two identical latest expected receipt date for {}".format(
                        item_no
                ))

        tmp_dict = {}
        
        # only one entry in the file, just pick it
        if len(expect_latest_index_list) == 1 and len(promise_latest_index_list) == 0:
            tmp_dict.update({po_no_list[expect_latest_index_list[0]]:top_dict[po_no_list[expect_latest_index_list[0]]]})

        elif len(expect_latest_index_list) == 0 and len(promise_latest_index_list) == 1:
            tmp_dict.update({po_no_list[promise_latest_index_list[0]]:top_dict[po_no_list[promise_latest_index_list[0]]]})
        

        elif len(expect_latest_index_list) > 1:
            # the date were found in same PO
            tmp_dict.update({po_no_list[expect_latest_index_list[0]]:top_dict[po_no_list[expect_latest_index_list[0]]]})
            tmp_dict.update({po_no_list[expect_latest_index_list[1]]:top_dict[po_no_list[expect_latest_index_list[1]]]})

        if tmp_dict:
            purchase_dict.update({item_no:tmp_dict})            

    with open('Purchase Lines_step_2.json', "w") as json_file:
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

    bom_dict = Dict_step_2['BOM list_step_1.json']
    filtered_dict = {}

    for key, value in Dict_step_2['customer order list_step_1.json'].items():
        customer_document_no, bom_no = key.split('---')
        customer_dict = value
        if bom_no in bom_dict:
            # print("there is associatekey in BOM for {}".format(value['Document No.']))
            bom_items_dict = bom_dict.get(bom_no, 'not found')

            for bom_item_no, detail_dict in bom_items_dict.items():

                value.update({bom_item_no:detail_dict})

            filtered_dict.update({key:value})
        else:
            print("customer order list err: {} no associate entry in BOM".format(bom_no))

    with open('customer order list_step_2.json', "w") as json_file:
        json.dump(filtered_dict, json_file, indent = 4) 


# use bom_no_id in the BOM List as the key to reorganize the customer order list to gather all customer order which belong to
# one bom item
def step_3_processing():

    with open('customer order list_step_2.json') as json_file:
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


    with open('Item List_step_1.json') as json_file:
        item_dict = json.load(json_file)

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
    
    with open('customer order list_step_3.json', "w") as json_file:
        json.dump(final_dict, json_file, indent = 4)     

# print("now is {}".format(date_time.strftime(FMT)))

def write_to_xls_file():
    final_order_dict = OrderedDict()
    with open('customer order list_step_3.json') as json_file:
        final_order_dict = json.load(json_file) 

    book = openpyxl.Workbook()
    sheet = book.active


    current_day = datetime.today() - timedelta(days=3)


    row_1_list = [  "Vendor",
                    'Production Code',
                    "Duty Class",
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

    header_font = Font(color="FF0000")
    sub_row_font = [
                        Font(color="eeaabb"),
                        Font(color="bb7766"),
                        Font(color="dd9988"),
                        Font(color="cc8877"),
                        Font(color="bb7766"),
                        Font(color="eeaabb"),
                        Font(color="bb7766"),
                        Font(color="eeaabb"),
                    ]


    for offset in range(14):
        next_day = current_day + timedelta(days=offset+1)
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
        sheet.cell(row=row_cnt, column=column_cnt+1).fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type = "solid")

    sheet.freeze_panes = "A2"

    sheet.print_title_rows='1:1'

    row_cnt += 1
    column_cnt = 1

    for bom_key, top_dict in final_order_dict.items():

        if top_dict.get("Duty Class", "unknown") in Interest_Duty_List:

            demand_dict = defaultdict(lambda: float(0))
            demand_dict_tmp = defaultdict(lambda: float(0))
            for middle_key, middle_dict in top_dict.items():
                if isinstance(middle_dict, dict):
                    per_unit_production_quantity = float(middle_dict["Production Quantity"])
                    scrap = float(middle_dict["Scrap %"])
                    order_quantity = float(middle_dict["Customer order Quantity"])
                    demand_quantity = order_quantity * per_unit_production_quantity * (1.0 + scrap * 0.01)

                    # till now only occur once
                    if middle_dict["Shipment Date"] not in demand_dict:
                        demand_dict_tmp.update({middle_dict["Shipment Date"]: demand_quantity})
                    # already have one, accumulate the demand on it
                    else:
                        demand_sum = demand_dict[middle_dict["Shipment Date"]] + demand_quantity
                        demand_dict_tmp.update({middle_dict["Shipment Date"]: demand_sum})

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

                sheet.cell(row=row_cnt, column=4).value = specific_column_list[sub_row]
                sheet.cell(row=row_cnt, column=4).font = sub_row_font[sub_row]

                # add demand number into the excel
                if 'Demand' in specific_column_list[sub_row]:
                    # put demand value for each following 14 days
                    for column_no in range(4, len(row_1_list)):
                        if row_1_list[column_no] in demand_dict:
                            sheet.cell(row=row_cnt, column=column_no + 1).value = demand_dict[row_1_list[column_no]]

                elif 'On Hand (W1)' in specific_column_list[sub_row]:
                    sheet.cell(row=row_cnt, column=5).value = int(top_dict.get("W1", 0))

                elif 'On Hand (W2)' in specific_column_list[sub_row]:
                    sheet.cell(row=row_cnt, column=5).value = int(top_dict.get("W2", 0))

                row_cnt += 1
                column_cnt = 1
        
        column_cnt = 1

    print("write {} row".format(row_cnt))

                
    book.save("sample.xlsx")


# read_data_from_files(file_dict)
step_2_processsing()
step_3_processing()
write_to_xls_file()

