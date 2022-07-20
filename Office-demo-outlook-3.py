## importing the necessary modules
from concurrent.futures import Executor
import os
from numpy import reciprocal
os.system("pip install pywin32 pandas")
import win32com.client as win32 # for using COM model for MS Office 365 application
import pandas as pd # to read and manipulate the excel sheet

#workbook=r"<file path>"
workbook=r"C:\Users\emaienj\Downloads\Book1.xlsx"
xls=pd.ExcelFile(workbook)
data_frame=pd.read_excel(xls,'Sheet1')

# for reading multiple xcel sheet use: data_frame=pd.read_excel(xls, sheet_name=[0,1,2,........])
# for reading multiple xcel sheet by sheet name use: data_frame=pd.read_excel(xls, sheet_name=['<sheet_name>','<sheet_name>',......])
# for reading all excel sheet use: data_frame=pd.read_excel(xls, sheet_name=None)

print(data_frame.keys())
recipients=[]
recipientsCC=[]
os.system("cls")
circle=input("Enter the Circle: ")
circles=[]
cr_detail=[]
cr_title=[]
executor=[]
for i in range(0,len(data_frame)):
        if data_frame.iloc[i]['Circle'] == circle:
             cr_detail.append(data_frame.iloc[i]['Cr Detail'])
             circles.append(data_frame.iloc[i]['Circle'])
             cr_title.append(data_frame.iloc[i]['Cr title'])
             executor.append(data_frame.iloc[i]['Executor'])
             recipients=data_frame.iloc[i]['Recipeint'].split(',') #adding recipients
             recipientsCC=data_frame.iloc[i]['Recipient in copy'].split(',') # adding CC
df = pd.DataFrame({'Cr Detail' :cr_detail,'Cr title': cr_title, 'Circle' : circles, 'Executor': executor})
df.reset_index(drop=True, inplace=True)
recipients=";".join(recipients)
recipientsCC=";".join(recipientsCC)

# Creating the COM outlook object to send the mail
outlook=win32.Dispatch('Outlook.Application')
msg=outlook.CreateItem(0) # 0 for outlook mail
html_body="""
<html>
<link rel="stylesheet" href="df_style.css">
<body>
<div><p>Hi team,</p><br><br>

<p>Please confirm below points so that we will approve CR’s.<br><br></p>

<p>1)  End nodes and service details are required which are running on respective MPBN device (in case of changes on Core/PACO/HLR devices ).<br></p>
<p>2)  Design Maker & Checker confirmation mail need to be shared for all planned activity on Core/PACO/HLR devices.<br></p>
<p>3)  KPI & Tester details need to be shared for all impacted nodes in Level-1 CR’s (SA).Also same details need to be shared for all Level-2 CR’s (NSA) with respect to changes on Core/PACO/HLR devices.<br><br></p>
</div>
<div>
    <p>{}</p>
</div>
<div>
    <p>Regards</p>
    <p>{}</p>
    <p>only for testing From Enjoy Maity</p>
    <p></p>
</div>
</body>
</html>
"""
sender = data_frame.iloc[0]['Executor']
msg.Subject="ONLY FOR TEST :Connected End Nodes and their services on MPBN devices: {}".format(circle)
msg.To= recipients
msg.CC= recipientsCC
#df.style()
msg.HTMLBody=html_body.format(df.to_html(index=False),sender)
msg.Save()
msg.Send()

#os.system("cls")
#print(len(data_frame))
#print(data_frame.iloc[0]['Circle'])

