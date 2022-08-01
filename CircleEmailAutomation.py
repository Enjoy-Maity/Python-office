class EmptyString (Exception):
    def __init__(self,msg):
        self.msg=msg

class ContainsInteger(Exception):
    def __init__(self,msg):
        self.msg=msg


#####################################################################
#############################    Sendmail   #########################
#####################################################################

def sendmail(dataframe,to,cc,body,subject,sender):
    outlook_mailer=win32.Dispatch('Outlook.Application')
    msg=outlook_mailer.CreateItem(0)
    html_body=body
    msg.Subject=subject
    msg.To=to
    msg.CC=cc
    dataframe=dataframe.style.set_table_styles([
        {'selector':'th','props':'border:1px solid black; color:white; background-color:rgb(0, 51, 204);text-align:center;'},
        {'selector':'tr','props':'border:1px solid black;text-align:center;'},
        {'selector':'td','props':'border:1px solid black;text-align:center;'},
        {'selector':'tr:nth-child(even)','props':'border:1px solid black;text-align:center;'}])
    dataframe=dataframe.hide(axis='index')
    msg.HTMLBody=html_body.format(dataframe.to_html(index=False),sender)
    msg.Save()
    msg.Send()

#####################################################################
############################# Fetch-details #########################
#####################################################################

def fetch_details(sender):
    user=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

    workbook=r"C:\Users\{}\Daily\MPBN Daily Planning Sheet.xlsx".format(user)
    excel=pd.ExcelFile(workbook)
    daily_plan_sheet=pd.read_excel(excel,'Planning Sheet')
    Email_ID=pd.read_excel(excel,'Mail Id')

    for j in range(0,len(daily_plan_sheet)):
        temp=daily_plan_sheet.at[j,'Circle']
        daily_plan_sheet.at[j,'Circle']=temp.upper()

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

            tomorrow=datetime.now()+timedelta(1)

            if daily_plan_sheet.iloc[j]['Circle']==circles[i]: # Adding constraint to check for CRs for next date only

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

        cir=circles[i]

        if cir=='DL':
            row_to_fetch=0

        elif cir=='PB':
            row_to_fetch=1

        elif cir=='HRY':
            row_to_fetch=2
        else :
            pass


        to=Email_ID.iloc[row_to_fetch]['To Mail List']
        cc=Email_ID.iloc[row_to_fetch]['Copy Mail List']
        
        subject=f"ONLY FOR TEST :Connected End Nodes and their services on MPBN devices: {cir}"
        body="""
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
            """
        sendmail(dataframe,to,cc,body,subject,sender)
        print(f"\nMail Sent for the Circle {cir}")
        time.sleep(5)
            


#####################################################################
################################# Main ##############################
#####################################################################   

if __name__=="__main__":
    
    try:
        import sys
        import os
        import re
        import time
        from datetime import datetime,timedelta
        import subprocess
        subprocess.run(["python.exe" ,"-m ","pip" ,"install" ,"--upgrade ","pip"],shell=True)
        import pkg_resources
        import pandas as pd
        import win32com.client as win32
        from openpyxl import load_workbook
        from openpyxl import Workbook
        import xlsxwriter
        import numpy
        sender=input("Enter your name to start the program or n/N to exit : ")
        if len(sender)==0:
            raise EmptyString("\nPlease enter your name not an Empty String.\n")
        
        elif (len(re.findall("\d",sender)))>0:
            raise ContainsInteger("\nInvalid Name as it contains Integer\n")
        
        elif sender=='n' or sender=='N':
            sys.exit(0) # exiting the program
        else:
            fetch_details(sender)

    except ModuleNotFoundError:
         required_modules={'pandas','pywin32','Jinja2', 'openpyxl','numpy','xlsxwriter'}
         installed_modules={pkg.key for pkg in pkg_resources.working_set}
         missing_modules= required_modules-installed_modules
         if missing_modules:
            python=sys.executable
            subprocess.check_call([python, '-m', 'pip','install',*missing_modules], stdout=subprocess.DEVNULL)
            print("Some Important Modules were absent and are now installed starting the program now.........")
        
            current_file=__file__ # gets the value of current running file
            subprocess.run(['python', current_file])
            sys.exit(0)

    except FileNotFoundError:
        working_directory=r"C:\Users\{}\Daily".format(subprocess.getoutput("echo %username%"))
        print("Check {} for MPBN Daily Planning Sheet.xlsx".format(working_directory))
    
    except ValueError:
         working_directory=r"C:\Users\{}\Daily".format(subprocess.getoutput("echo %username%"))
         print("Check {} for MPBN Daily Planning Sheet.xlsx for all the requirement sheet".format(working_directory))
    
    except EmptyString as error:
        print(error)
        current_file=__file__ # gets the value of current running file
        subprocess.run(['python', current_file])
        sys.exit(0)
    
    except ContainsInteger as error:
        print(error)
        current_file=__file__ # gets the value of current running file
        subprocess.run(['python', current_file])
        sys.exit(0)