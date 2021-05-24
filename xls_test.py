import openpyxl
from datetime import datetime, timedelta
import pprint

FMT = '%Y-%m-%d %H:%M:%S'
FMT_1 = '%Y-%m-%d, %H:%M:%S'

time = datetime.strptime("2021-02-13 00:00:00", FMT)

date_time = datetime.now()

print("now is {}".format(date_time.strftime(FMT)))

wb= openpyxl.load_workbook('BR-May-20-2021a-mod-1.xlsx')
names = wb.sheetnames
print(wb.sheetnames)
sheet=wb["Buyer's Report"]
colum_dict = {row_5 : {}}

for x in range (5,6):
    date_count = 0

    for y in range(1,50):

        if isinstance(sheet.cell(row=x,column=y).value, datetime):
            if date_count == 0:
                colum_name_dict.update({y:time})
            elif date_count < 4:
                time_str_list = time.strftime(FMT).split('-')
                time_str_list[1] = str(int(time_str_list[1]) + date_count)
                new_time = "-".join(time_str_list)
                colum_name_dict.update({y:datetime.strptime(new_time, FMT)})

            elif date_count == 4 or date_count == 5:
                time_str_list = time.strftime(FMT).split('-')
                time_str_list[1] = str(int(time_str_list[1]) + date_count + 1).zfill(2)
                time_str_list[2] = "01 00:00:00"
                new_time = "-".join(time_str_list)
                colum_name_dict.update({y:datetime.strptime(new_time, FMT) - timedelta(days=1)})     

            elif date_count > 5 and date_count < 18:
                start_time = colum_name_dict[18]
                colum_name_dict.update({y:start_time + timedelta(days=date_count - 6 + 7)})    
 
            elif date_count >= 18:
                start_time = colum_name_dict[41]
                colum_name_dict.update({y:start_time + timedelta(days=(date_count - 17) * 7)})  
            date_count += 1

        else:
            colum_name_dict.update({y:sheet.cell(row=x,column=y).value})

pp = pprint.PrettyPrinter(indent=4)
pp.pprint(colum_name_dict)

