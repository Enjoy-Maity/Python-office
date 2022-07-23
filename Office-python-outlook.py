from msilib.schema import Directory


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

try:
    import pandas as pd
    import sys
    import os
    import subprocess
    import pkg_resources
    import win32com.client as win32
    from openpyxl import load_workbook

    #Creating the function for 
    def sendmail():
        outlook_mailer=win32.Dispatch('Outlook.Application')
        msg=outlook_mailer.CreateItem(0)
        html_body='''
        <html>
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
        print(circles)
        #Mail_id=pd.ExcelFile(Workbook,'Mail Id')
        #sendmail()
    
    
except(ModuleNotFoundError,FileNotFoundError) as error:
    handle(error)

if __name__=="__main__":
    fetch_details()