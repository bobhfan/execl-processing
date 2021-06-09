from datetime import date as date_op, datetime, timedelta

day = datetime.now().date()
day_pre = datetime.now().date() - timedelta(days=1)

print(day.strftime("%m%d"))
print(day_pre.strftime("%m%d"))