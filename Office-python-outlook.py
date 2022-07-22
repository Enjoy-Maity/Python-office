try:
    import pandas as pd
    import sys
    import os
    import subprocess
    import pkg_resources
    import win32com.client as win32

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
        msg.To=
        msg.Save()
        msg.Send()

    def fetch_details():
        User=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 
        Workbook=r"C:\Users\{}\Daily".format(User)
        daily_plan_sheet=pd.ExcelFile(Workbook,'Planning Sheet')
        
    
except(ModuleNotFoundError,FileExistsError,FileNotFoundError):
    '''
    checking and then installing the important modules
    '''
    required_modules={'pandas','pywin32','Jinja2'}
    installed_modules={pkg.key for pkg in pkg_resources.working_set}
    missing_modules= required_modules-installed_modules

    if missing_modules:
        python=sys.executable
        subprocess.check_call([python, '-m', 'pip','install',*missing_modules], stdout=subprocess.DEVNULL)