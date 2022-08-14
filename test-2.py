import subprocess
import pandas as pd
from datetime import timedelta,datetime
tomorrow=datetime.now()+timedelta(1)
user=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

workbook=r"C:\Users\{}\Daily\MPBN Daily Planning Sheet.xlsx".format(user)
excel=pd.ExcelFile(workbook)
daily_plan_sheet=pd.read_excel(excel,'Planning Sheet')
daily_plan_sheet.fillna("NA",inplace=True)
# for j in range(0,len(daily_plan_sheet)):

#             #tomorrow=datetime.now()+timedelta(1)
#             #print(str(tomorrow.strftime("%d-%m-%Y")))
#     temp=daily_plan_sheet.iloc[j]['Execution Date']
#     if temp==tomorrow.strftime("%Y/%m/%d"):
#         print("Yes")
#     else:
#         print("No")

print(len(daily_plan_sheet[daily_plan_sheet['Execution Date']==tomorrow.strftime("%Y-%m-%d")]))
for j in range(0,len(daily_plan_sheet[daily_plan_sheet['Execution Date']==tomorrow.strftime("%Y-%m-%d")])):
    print(daily_plan_sheet[daily_plan_sheet['Execution Date']==tomorrow.strftime("%Y-%m-%d")])