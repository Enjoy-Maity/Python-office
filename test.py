from openpyxl import Workbook
import pandas as pd
import subprocess
#from datetime import datetime,timedelta

user=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

workbook=r"C:\Users\{}\Daily\MPBN Daily Planning Sheet.xlsx".format(user)
print(workbook)
excel=pd.ExcelFile(workbook)
daily_plan_sheet=pd.read_excel(excel,'Planning Sheet')
print(len(daily_plan_sheet))
Email_ID=pd.read_excel(excel,'Mail Id')
print(len(Email_ID))

for j in range(0,len(daily_plan_sheet)):
    str=daily_plan_sheet.at[j,'Circle']
    daily_plan_sheet.at[j,'Circle']=str.upper()

circles=daily_plan_sheet['Circle'].unique()
print(circles)
email_id_list=Email_ID['Circle'].unique()
print(email_id_list)
# print(circles) # checking for all the unique values of circles in the MPBN Planning Sheets
remainder=list(set(circles)-set(email_id_list))
remainder_list=",".join(remainder)
if len(remainder)>0:
    print(f"\nMail could not be sent for {remainder_list} as there's no email id present for the {remainder_list} in the Email ID sheet in MPBN Daily Planning Sheet")

circles=list(set(circles)-set(remainder))
    
for i in range(0,len(circles)):

    execution_date=[]       #  list for collecting execution date of each Cr
    circle=[]               #  list for collecting circle of each CR
    maintenance_window=[]   #  list for collecting the maintenance window of each CR
    cr_no=[]                #  list for collecting the CR No
    activity_title=[]       #  list for collecting the activity title each CR
    risk=[]                 #  list for collecting the risk level of each CR
    location=[]             #  list for collecting the location of each CR

    for j in range(0,len(daily_plan_sheet)):

        #tomorrow=datetime.now()+timedelta(1)

        if  daily_plan_sheet.iloc[j]['Circle']==circles[i]: # Adding constraint to check for CRs for next date only

            execution_date.append(daily_plan_sheet.iloc[j]['Execution Date'])
            maintenance_window.append(daily_plan_sheet.iloc[j]['Maintenance Window'])
            cr_no.append(daily_plan_sheet.iloc[j]['CR NO'])
            activity_title.append(daily_plan_sheet.iloc[j]['Activity Title'])
            risk.append(daily_plan_sheet.iloc[j]['Risk'])
            circle.append(daily_plan_sheet.iloc[j]['Circle'])
            location.append(daily_plan_sheet.iloc[j]['Location'])

            dictionary_for_insertion={'Execution Date':execution_date, 'Maintenance Window':maintenance_window, 'CR NO':cr_no, 'Activity Title':activity_title, 'Risk':risk,'Location':location,'Circle':circle}
            dataframe=pd.DataFrame(dictionary_for_insertion)
            dataframe.reset_index(drop=True,inplace=True)
            
            print(dataframe)
