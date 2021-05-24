import openpyxl
from datetime import datetime, timedelta
import pprint
import json
from collections import OrderedDict

FMT = '%Y-%m-%d'


# time = datetime.strptime("2021-02-13 00:00:00", FMT)

# date_time = datetime.now()

# print("now is {}".format(date_time.strftime(FMT)))
file_dict = {
                'BOM list.xlsx' : 1,
                'Item List.xlsx' : 2,
                'customer order list.xlsx': 1,
            }


Item_List_Interest_Field_List = ['Duty Class', 
                                'W1',
                                'W2',
                                'Base Unit of Measure' ,
                                'Vendor No.',
                                ]

Interest_Duty_List = ["FILM", "BAG",  ]

Dict_step_2 = {
                'BOM list_step_1.json' : {},
                'Item List_step_1.json' : {},
                'customer order list_step_1.json': {}
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
                        
                        item_no_value = sheet.cell(row=line, column=colum_n).value

                    elif field_name_dict[colum_n] in Item_List_Interest_Field_List:
                        tmp_dict.update({field_name_dict[colum_n] : sheet.cell(row=line, column=colum_n).value}) 

                    elif 'Blocked' in field_name_dict[colum_n]:
                        item_block = sheet.cell(row=line, column=colum_n).value

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
                    data_dict.update({item_no_value : tmp_dict})

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
    
    for document_no, value in Dict_step_2['BOM list_step_1.json'].items():
        tmp_dict = {}
        for bom_no, detail_dict in value.items():
            if bom_no not in Dict_step_2['Item List_step_1.json']:
                print("item_list err: no record for {}".format(bom_no))
            elif not Dict_step_2['Item List_step_1.json'][bom_no]['Duty Class']:
                print("item_list err: empty duty class for {}".format(bom_no))
            else:
                for key, value in Dict_step_2['Item List_step_1.json'][bom_no].items():
                    detail_dict.update({key : value})

                tmp_dict.update({bom_no:detail_dict})
        
        bom_dict.update({document_no:tmp_dict})
    
    with open('BOM list_step_2.json', "w") as json_file:
        json.dump(bom_dict, json_file, indent = 4)    

    filtered_dict = {}

    for key, value in Dict_step_2['customer order list_step_1.json'].items():
        customer_document_no, bom_no = key.split('---')
        customer_dict = value
        if bom_no in bom_dict:
            # print("there is associatekey in BOM for {}".format(value['Document No.']))
            bom_items_dict = bom_dict.get(bom_no, 'not found')

            for bom_item_no, detail_dict in bom_items_dict.items():

                if detail_dict['Duty Class'] in Interest_Duty_List:

                    # print('found {} in the scope '.format(detail_dict['duty class']))

                    value.update({bom_item_no:detail_dict})

            filtered_dict.update({key:value})
        else:
            print("customer order list err: {} no associate entry in BOM".format(bom_no))

    with open('customer order list_step_2.json', "w") as json_file:
        json.dump(filtered_dict, json_file, indent = 4) 

    bag_list = []
    film_list = []

    for document_no, top_dict in filtered_dict.items():
        for bom_no, level_2_dict in top_dict.items():
            if isinstance(level_2_dict, dict):
                if 'FILM' in level_2_dict['Duty Class']:
                    film_list.append(bom_no)
                elif 'BAG' in level_2_dict['Duty Class']:
                    bag_list.append(bom_no)
    
    bag_set = set(bag_list)
    film_set = set(film_list)

    print("There are {} items of film in {} kinds of FILM and {} items of bag in {} kinds of BAG". \
            format(len(film_list), len(film_set), len(bag_list), len(bag_set)))

# use bom_no_id in the BOM List as the key to reorganize the customer order list to gather all customer order which belong to
# one bom item
def step_3_processing():

    with open('customer order list_step_2.json') as json_file:
        custom_dict = json.load(json_file)

    final_dict = {}

    for custom_no, top_dict in custom_dict.items():
        tmp_dict = {}
        # first iteration: found all non_dict pairs to have temp dict to update them into final dict
        for key, value in top_dict.items():
            if not isinstance(value, dict):
                tmp_dict.update({key:value})

        # second iteration: finding out all bom material
        for item_no, level_2_dict in top_dict.items():
            if isinstance(level_2_dict, dict):
                real_key = item_no

                # if materail not exist, create empty dict for adding
                if real_key not in final_dict:
                    final_dict.update({real_key:{}})

                # add bom detail info into dict
                for key, value in level_2_dict.items():
                    final_dict[real_key].update({key:value})

                # after add bom info into the dict, should append customer order info
                final_dict[real_key].update({custom_no:tmp_dict})

    total_cnt = 0
    for bom_key, top_dict in final_dict.items():
        order_cnt = 0
        for item_no, level_2_dict in top_dict.items():
            if isinstance(level_2_dict, dict):
                order_cnt += 1
                total_cnt += 1

        print("{} of duty '{}' showd up in {} customer orders".format(bom_key, top_dict['Duty Class'], order_cnt))
    
    with open('customer order list_step_3.json', "w") as json_file:
        json.dump(final_dict, json_file, indent = 4)     


def write_to_xls_file():
    final_order_dict = OrderedDict()
    with open('customer order list_step_3.json') as json_file:
        final_order_dict = json.load(json_file) 

    book = openpyxl.Workbook()
    sheet = book.active

    row_1_list = [  'No.',
                    "Description FG",
                    "Unit of Measure Code",
                    "Quantity",
                    "Scrap %",
                    "W1",
                    "W2",
                    "Duty Class",
                    "Vendor No.",
                    "Customer Order No.",
                    "Location Code",
                    "Customer order Quantity",
                    "Unit of Measure Code",
                    "Shipment Date",
                    "Outstanding Quantity",
                ]

    row_cnt = 1
    column_cnt = 1

    for column_cnt in range(len(row_1_list)):
        sheet.cell(row=row_cnt, column=column_cnt+1).value = row_1_list[column_cnt]

    row_cnt += 1
    column_cnt = 1

    for bom_key, top_dict in final_order_dict.items():
        sheet.cell(row=row_cnt, column=column_cnt).value = bom_key
        sheet.cell(row=row_cnt, column=2).value = top_dict["Description FG"]
        sheet.cell(row=row_cnt, column=3).value = top_dict["Unit of Measure Code"]
        sheet.cell(row=row_cnt, column=4).value = top_dict["Quantity"]
        sheet.cell(row=row_cnt, column=5).value = top_dict["Scrap %"]
        sheet.cell(row=row_cnt, column=6).value = top_dict["W1"]
        sheet.cell(row=row_cnt, column=7).value = top_dict["W2"]
        sheet.cell(row=row_cnt, column=8).value = top_dict["Duty Class"]
        sheet.cell(row=row_cnt, column=9).value = top_dict["Vendor No."] 
        for middle_key, middle_dict in top_dict.items():
            if isinstance(middle_dict, dict):
                sheet.cell(row=row_cnt, column=10).value = middle_key
                sheet.cell(row=row_cnt, column=11).value = middle_dict["Location Code"]
                sheet.cell(row=row_cnt, column=12).value = middle_dict["Customer order Quantity"]
                sheet.cell(row=row_cnt, column=13).value = middle_dict["Unit of Measure Code"]
                sheet.cell(row=row_cnt, column=14).value = middle_dict["Shipment Date"]
                sheet.cell(row=row_cnt, column=15).value = middle_dict["Outstanding Quantity"]

                row_cnt += 1
                column_cnt = 1

        column_cnt = 1

    print("write {} row".format(row_cnt))

                
    book.save("sample.xlsx")


read_data_from_files(file_dict)
step_2_processsing()
step_3_processing()
write_to_xls_file()

