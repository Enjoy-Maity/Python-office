#####################################################################
######################### Custom Exceptions #########################
#####################################################################

class EmptyString (Exception):
    def __init__(self,msg):
        self.msg=msg

class ContainsInteger(Exception):
    def __init__(self,msg):
        self.msg=msg


#####################################################################
#############################    Sendmail   #########################
#####################################################################

def sendmail(dataframe,to,cc,circle,sender):
    outlook_mailer=win32.Dispatch('Outlook.Application')
    msg=outlook_mailer.CreateItem(0)
    html_body="""
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
    msg.Subject="ONLY FOR TEST :Connected End Nodes and their services on MPBN devices: {}".format(circle)
    msg.To=to
    msg.CC=cc
    dataframe=dataframe.style.set_table_styles([
        {'selector':'th','props':'border:1px solid black;'},
        {'selector':'tr','props':'border:1px solid black;'},
        {'selector':'td','props':'border:1px solid black;'},
        {'selector':'tr:nth-child(even)','props':'border:1px solid black;'}])
    dataframe=dataframe.hide(axis='index')
    msg.HTMLBody=html_body.format(dataframe.to_html(index=False),sender)
    msg.Save()
    msg.Send()

#####################################################################
#############################  Paco_cscore  #########################
#####################################################################

def paco_cscore(Workbook):
    excel=pd.ExcelFile(Workbook)
    daily_plan_sheet=pd.read_excel(excel,'Planning Sheet')
    daily_plan_sheet.fillna("Not Available")
    Email_Id=pd.read_excel(Workbook,'Mail Id')
    #print(daily_plan_sheet)
    sheetname="PS Core-Inter Domain"
    category="MPBN-MS"
    owner_domain="SRF MPBN"
    team_leader="Karan Loomba"
    Cr_no=[]
    executor=[]
    validator=[]
    impact=[]
    circle=[]
    activity_title=[]
    cr_owner_domain=[]
    inter_domain=[]
    cr_category=[]
    node=[]
    Kpis=[]
    for i in range(0,len(daily_plan_sheet)):
        if daily_plan_sheet.iloc[i]['Domain kpi']=="PS Core" or daily_plan_sheet.iloc[i]['Domain kpi']=="Paco-circle":
            Cr_no.append(daily_plan_sheet.iloc[i]['CR NO'])
            cr_category.append(category)
            impact.append(daily_plan_sheet.iloc[i]['Impact'])
            circle.append(daily_plan_sheet.iloc[i]['Circle'])
            activity_title.append(daily_plan_sheet.iloc[i]['Activity Title'])
            cr_owner_domain.append(owner_domain)
            executor.append(daily_plan_sheet.iloc[i]['Change Responsible'])
            technical_validator=daily_plan_sheet.iloc[i]['Technical Validator']
            if technical_validator==team_leader:
                validator.append(team_leader)
            else:
                tech_validator_team_leader=technical_validator+"/"+team_leader
                validator.append(tech_validator_team_leader)
            inter_domain.append(daily_plan_sheet.iloc[i]['Domain kpi'])
            node.append(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
            Kpis.append(daily_plan_sheet.iloc[i]['KPI DETAILS'])

    dictionary1={'CR':Cr_no,'CR Category':cr_category,'Impact*':impact,'Circle':circle,'Activity Title':activity_title,'CR Owner Domain':cr_owner_domain,'Executor':executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Node':node,'KPIs':Kpis}
    df=pd.DataFrame(dictionary1)
    df.reset_index(drop=True,inplace=True)
    print(df)

    """
    # this is one way of writing into the exccel sheet but it didn't work

    excel_wb=win32.gencache.EnsureDispatch("Excel.Application")
    wb=excel_wb.Workbooks.Open(Workbook)
    #for i in range()
    #print(wb.Sheets(1).Name)
    
    ws=wb.Sheets(sheetname)
    used=ws.UsedRange
    print(used.Row+used.Rows.Count-2)
    startrow=used.Row+used.Rows.Count-1
    startcol=1
    ws=wb.Worksheets(sheetname)
    ws.Range(ws.Cells(startrow,startcol),ws.Cells(startrow+len(df.index)-1,startcol+len(df.columns)-1)).Value=df.values
    wb.SaveAs("MPBN Daily Planning Sheet")
    wb.Close()
    
    """
    """
    # another alternate way to write excel but it didn't work
    writer=pd.ExcelWriter(Workbook,mode='a',engine='openpyxl',if_sheet_exists='replace')
    df.to_excel(writer,sheet_name=sheetname,index=False)
    writer.save()
    """
    """
    sheet_mapping={sheetname:df}
    with xw.App(visible=False) as app:
        wb=app.books.open(Workbook)

        current_sheets =[sheet.name for sheet in wb.sheets]
        print(wb.sheets.count)

        for sheet_name in sheet_mapping.keys():
            if sheet_name in current_sheets:
                wb.sheets(sheet_name).range('A1').value = sheet_mapping.get(sheet_name)    
            
            else:
                new_sheet=wb.sheets.add(after=wb.sheets.count)
                new_sheet.range('A1').value = sheet_mapping.get(sheet_name)
                new_sheet.name=sheet_name
            
        wb.save()
        wb.close()
    """
    
    writer=pd.ExcelWriter(Workbook,engine='xlsxwriter')
    daily_plan_sheet.to_excel(writer,sheet_name='Planning Sheet',index=False)
    df.to_excel(writer,sheet_name=sheetname,index=False)
    Email_Id.to_excel(writer,sheet_name='Mail Id',index=False)

    writer.save()

#####################################################################
############################# Fetch-details #########################
#####################################################################

def fetch_details(sender):
    User=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

    Workbook=r"C:\Users\{}\Daily\MPBN Daily Planning Sheet.xlsx".format(User)
    """excel=pd.ExcelFile(Workbook)
    daily_plan_sheet=pd.read_excel(excel,'Planning Sheet')
    Email_ID=pd.read_excel(excel,'Mail Id')

    if 'CS Core-Inter Domain' in wb.sheetnames:
        pass
    else :
        cscore_interdomain=wb.create_sheet()
        cscore_interdomain.title="CS Core-Inter Domain"

    circles=daily_plan_sheet['Circle'].unique()
    # print(circles) # checking for all the unique values of circles in the MPBN Planning Sheets
        
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
        for k in range(0,len(Email_ID)):
            if(Email_ID.iloc[k]['Circle']==circles[i]):
                to=Email_ID.iloc[k]['To Mail List']
                cc=Email_ID.iloc[k]['Copy Mail List']
                cir=circles[i]
        sendmail(dataframe,to,cc,cir,sender)"""
    paco_cscore(Workbook)


#####################################################################
################################# Main ##############################
#####################################################################   

if __name__=="__main__":
    
    try:
        import sys
        import os
        import re
        import subprocess
        subprocess.run(["python.exe" ,"-m ","pip" ,"install" ,"--upgrade ","pip"],shell=True)
        import pkg_resources
        import pandas as pd
        import win32com.client as win32
        from openpyxl import load_workbook
        from openpyxl import Workbook
        import xlwings as xw
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
         required_modules={'pandas','pywin32','Jinja2', 'openpyxl','numpy','xlwings','xlsxwriter'}
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


   