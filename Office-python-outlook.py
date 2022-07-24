#from msilib.schema import Directory


def handle(error):
    if error== 'ModuleNotFoundError':
        required_modules={'pandas','pywin32','Jinja2', 'openpyxl','numpy'}
        installed_modules={pkg.key for pkg in pkg_resources.working_set}
        missing_modules= required_modules-installed_modules
        if missing_modules:
            python=sys.executable
            subprocess.check_call([python, '-m', 'pip','install',*missing_modules], stdout=subprocess.DEVNULL)
        print("Some Important Modules were absent and are now installed starting the program now.........")
        
        current_file=__file__ # gets the value of current running file
        subprocess.run(['python', current_file])
        sys.exit(0)
    elif error=='FileNotFoundError':
        working_directory=r"C:\Users\{}\Daily".format(subprocess.getoutput("echo %username%"))
        print("Check {} for MPBN Daily Planning Sheet.xlsx".format(working_directory))
    elif error=='ValueError':
        working_directory=r"C:\Users\{}\Daily".format(subprocess.getoutput("echo %username%"))
        print("Check {} for MPBN Daily Planning Sheet.xlsx for all the requirement sheet".format(working_directory))

try:
    import pandas as pd
    import sys
    import os
    import subprocess
    import pkg_resources
    import win32com.client as win32
    from openpyxl import load_workbook
    import numpy

    #Creating the function for 
    def sendmail():
        outlook_mailer=win32.Dispatch('Outlook.Application')
        msg=outlook_mailer.CreateItem(0)
        html_body='''
        <html>
            <style>
                .mystyle {
                        font-size: 11pt; 
                        font-family: Arial;
                        border-collapse: collapse; 
                        border: 1px solid black;

                    }

                .mystyle td, th {
                        padding: 5px;
                    }

                .mystyle tr:nth-child(even) {
                        background: #E0E0E0;
                    }

                .mystyle tr:hover {
                        background: silver;
                        cursor: pointer;
                    }
            </style>
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
        '''
        msg.HTMLBody=html_body.format()
        #msg.To=
        msg.Save()
        msg.Send()
    
    def fetch_details():
        User=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

        Workbook=r"C:\Users\{}\Daily\MPBN Daily Planning Sheet.xlsx".format(User)
        wb=load_workbook(Workbook,read_only=False)
        excel=pd.ExcelFile(Workbook)
        daily_plan_sheet=pd.read_excel(excel,'Planning Sheet')
        
        if 'PS Core-Inter Domain' in wb.sheetnames:
            pscore_interdomain=pd.read_excel(Workbook,'PS Core-Inter Domain')
        else:
           pscore_interdomain=wb.create_sheet()
           pscore_interdomain.title="PS Core-Inter Domain"
        
        if 'CS Core-Inter Domain' in wb.sheetnames:
            cscore_interdomain=pd.read_excel(Workbook,'CS Core-Inter Domain')
        else :
            cscore_interdomain=wb.create_sheet()
            cscore_interdomain.title="CS Core-Inter Domain"

        circles=daily_plan_sheet['Circle'].unique()
        # print(circles) # checking for all the unique values of circles in the MPBN Planning Sheets
        print(type(circles))
        
        for i in range(0,len(circles)):
            execution_date=[]       #  list for collecting execution date of each Cr
            circle=[]               #  list for collecting circle of each CR
            maintenance_window=[]   #  list for collecting the maintenance window of each CR
            cr_no=[]                #  list for collecting the CR No
            activity_title=[]       #  list for collecting the activity title each CR
            risk=[]                 #  list for collecting the risk level of each CR


            for j in range(0,len(daily_plan_sheet)):
                if daily_plan_sheet.iloc[j]['Circle']==circles[i]:
                    execution_date.append(daily_plan_sheet.iloc[j]['Execution Date'])
                    maintenance_window.append(daily_plan_sheet.iloc[j]['Maintenance Window'])
                    cr_no.append(daily_plan_sheet.iloc[j]['CR NO'])
                    activity_title.append(daily_plan_sheet.iloc[j]['Activity Title'])
                    risk.append(daily_plan_sheet.iloc[j]['Risk'])
                    circle.append(daily_plan_sheet.iloc[j]['Circle'])
            dictionary_for_insertion={'Execution Date':execution_date, 'Maintenance Window':maintenance_window, 'CR NO':cr_no, 'Activity Title':activity_title, 'Risk':risk, 'Circle':circle}
            dataframe=pd.DataFrame(dictionary_for_insertion)
            dataframe.reset_index(drop=True,inplace= True)
            



           # sendmail()

        #sendmail()
    
    
except(ModuleNotFoundError,FileNotFoundError,ValueError) as error:
    handle(error)

if __name__=="__main__":
    fetch_details()