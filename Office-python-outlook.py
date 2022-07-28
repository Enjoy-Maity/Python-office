#####################################################################
######################### Custom Exceptions #########################
#####################################################################

from datetime import date


class EmptyString (Exception):
    def __init__(self,msg):
        self.msg=msg

class ContainsInteger(Exception):
    def __init__(self,msg):
        self.msg=msg

class EmailIDError(Exception):
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
#############################  Paco_cscore  #########################
#####################################################################

def paco_cscore(workbook,sender):
    

    daily_plan_sheet=pd.read_excel(workbook,'Planning Sheet')
    Email_Id=pd.read_excel(workbook,'Mail Id')
    #print(daily_plan_sheet)

    # Sheetnames
    sheetname="PS Core-Inter Domain"
    sheetname2="CS Core-Inter Domain"
    sheetname3="RAN-Inter Domain"


    category="MPBN-MS"
    owner_domain="SRF MPBN"
    team_leader="Karan Loomba"

    ####################################################### Entering details for ps core or paco circle ###########################################################
    execution_date=[]
    maintenance_window=[]
    mpbn_cr_no=[]
    location=[]
    mpbn_change_responsible_executor=[]
    validator=[]
    impact=[]
    circle=[]
    mpbn_activity_title=[]
    cr_owner_domain=[]
    inter_domain=[]
    cr_category=[]
    impacted_node_details=[]
    Kpis_to_be_monitored=[]
    # Execution Date	Maintenance Window	MPBN CR NO	CR Category	Impact*	Location	Circle	MPBN Activity Title	CR Owner Domain	MPBN Change Responsible	Technical Validator/Team Lead	InterDomain	Impacted Node Details	KPI's to be monitored
    for i in range(0,len(daily_plan_sheet)):
        if daily_plan_sheet.iloc[i]['Domain kpi']=="PS Core" or daily_plan_sheet.iloc[i]['Domain kpi']=="Paco-circle":
            execution_date.append(daily_plan_sheet.iloc[i]['Execution Date'])
            maintenance_window.append(daily_plan_sheet.iloc[i]['Maintenance Window'])
            mpbn_cr_no.append(daily_plan_sheet.iloc[i]['CR NO'])
            cr_category.append(category)
            impact.append(daily_plan_sheet.iloc[i]['Impact'])
            location.append(daily_plan_sheet.iloc[i]['Location'])
            txt=str(daily_plan_sheet.iloc[i]['Circle'])
            circle.append(txt.upper())
            mpbn_activity_title.append(daily_plan_sheet.iloc[i]['Activity Title'])
            cr_owner_domain.append(owner_domain)
            mpbn_change_responsible_executor.append(daily_plan_sheet.iloc[i]['Change Responsible'])
            technical_validator=daily_plan_sheet.iloc[i]['Technical Validator']
            if technical_validator==team_leader:
                validator.append(team_leader)
            else:
                tech_validator_team_leader=technical_validator+"/"+team_leader
                validator.append(tech_validator_team_leader)
            inter_domain.append(daily_plan_sheet.iloc[i]['Domain kpi'])
            impacted_node_details.append(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
            Kpis_to_be_monitored.append(daily_plan_sheet.iloc[i]['KPI DETAILS'])

    dictionary1={'CR':mpbn_cr_no,'Maintenance Window':maintenance_window,'CR Category':cr_category,'Impact*':impact,'Location':location,'Circle':circle,'MPBN Activity Title':mpbn_activity_title,'CR Owner Domain':cr_owner_domain,'Change Responsible':mpbn_change_responsible_executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Impacted Node Details':impacted_node_details,'KPIs to be monitored':Kpis_to_be_monitored}
    df=pd.DataFrame(dictionary1)


    ######################################################### Entering details for Cs core #######################################################################
    execution_date=[]
    maintenance_window=[]
    mpbn_cr_no=[]
    location=[]
    mpbn_change_responsible_executor=[]
    validator=[]
    impact=[]
    circle=[]
    mpbn_activity_title=[]
    cr_owner_domain=[]
    inter_domain=[]
    cr_category=[]
    impacted_node_details=[]
    Kpis_to_be_monitored=[]
    for i in range(0,len(daily_plan_sheet)):
        if daily_plan_sheet.iloc[i]['Domain kpi']=="CS Core":
            execution_date.append(daily_plan_sheet.iloc[i]['Execution Date'])
            maintenance_window.append(daily_plan_sheet.iloc[i]['Maintenance Window'])
            mpbn_cr_no.append(daily_plan_sheet.iloc[i]['CR NO'])
            cr_category.append(category)
            impact.append(daily_plan_sheet.iloc[i]['Impact'])
            location.append(daily_plan_sheet.iloc[i]['Location'])
            txt=str(daily_plan_sheet.iloc[i]['Circle'])
            circle.append(txt.upper())
            mpbn_activity_title.append(daily_plan_sheet.iloc[i]['Activity Title'])
            cr_owner_domain.append(owner_domain)
            mpbn_change_responsible_executor.append(daily_plan_sheet.iloc[i]['Change Responsible'])
            technical_validator=daily_plan_sheet.iloc[i]['Technical Validator']
            if technical_validator==team_leader:
                validator.append(team_leader)
            else:
                tech_validator_team_leader=technical_validator+"/"+team_leader
                validator.append(tech_validator_team_leader)
            inter_domain.append(daily_plan_sheet.iloc[i]['Domain kpi'])
            impacted_node_details.append(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
            Kpis_to_be_monitored.append(daily_plan_sheet.iloc[i]['KPI DETAILS'])
    dictionary2={'CR':mpbn_cr_no,'Maintenance Window':maintenance_window,'CR Category':cr_category,'Impact*':impact,'Location':location,'Circle':circle,'MPBN Activity Title':mpbn_activity_title,'CR Owner Domain':cr_owner_domain,'Change Responsible':mpbn_change_responsible_executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Impacted Node Details':impacted_node_details,'KPIs to be monitored':Kpis_to_be_monitored}
    df2=pd.DataFrame(dictionary2)


    ##########################################################  Entering details for RAN  ########################################################################
    execution_date=[]
    maintenance_window=[]
    mpbn_cr_no=[]
    location=[]
    mpbn_change_responsible_executor=[]
    validator=[]
    impact=[]
    circle=[]
    mpbn_activity_title=[]
    cr_owner_domain=[]
    inter_domain=[]
    cr_category=[]
    impacted_node_details=[]
    Kpis_to_be_monitored=[]
    oss_name=[]
    oss_IP=[]
    # Execution Date	Maintenance Window	MPBN CR NO	CR Category	Impact*	Location	Circle	MPBN Activity Title	CR Owner Domain	MPBN Change Responsible	Technical Validator/Team Lead	InterDomain	Impacted Node Details	KPI's to be monitored
    for i in range(0,len(daily_plan_sheet)):
        if daily_plan_sheet.iloc[i]['Domain kpi']=="RAN":
            execution_date.append(daily_plan_sheet.iloc[i]['Execution Date'])
            maintenance_window.append(daily_plan_sheet.iloc[i]['Maintenance Window'])
            mpbn_cr_no.append(daily_plan_sheet.iloc[i]['CR NO'])
            cr_category.append(category)
            impact.append(daily_plan_sheet.iloc[i]['Impact'])
            location.append(daily_plan_sheet.iloc[i]['Location'])
            txt=str(daily_plan_sheet.iloc[i]['Circle'])
            circle.append(txt.upper())
            mpbn_activity_title.append(daily_plan_sheet.iloc[i]['Activity Title'])
            cr_owner_domain.append(owner_domain)
            mpbn_change_responsible_executor.append(daily_plan_sheet.iloc[i]['Change Responsible'])
            technical_validator=daily_plan_sheet.iloc[i]['Technical Validator']
            if technical_validator==team_leader:
                validator.append(team_leader)
            else:
                tech_validator_team_leader=technical_validator+"/"+team_leader
                validator.append(tech_validator_team_leader)
            inter_domain.append(daily_plan_sheet.iloc[i]['Domain kpi'])
            impacted_node_details.append(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
            Kpis_to_be_monitored.append(daily_plan_sheet.iloc[i]['KPI DETAILS'])
            oss_name.append(daily_plan_sheet.iloc[i]['oss name'])
            oss_IP.append(daily_plan_sheet.iloc[i]['oss ip'])

    dictionary3={'CR':mpbn_cr_no,'Maintenance Window':maintenance_window,'CR Category':cr_category,'Impact*':impact,'Location':location,'Circle':circle,'MPBN Activity Title':mpbn_activity_title,'CR Owner Domain':cr_owner_domain,'Change Responsible':mpbn_change_responsible_executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Impacted Node Details':impacted_node_details,'KPIs to be monitored':Kpis_to_be_monitored,'OSS Name':oss_name,'OSS IP':oss_IP}
    df3=pd.DataFrame(dictionary3)


    df.reset_index(drop=True,inplace=True)
    df2.reset_index(drop=True,inplace=True)
    df3.reset_index(drop=True,inplace=True)
    list_of_interdomains=["CS Core","PS Core","RAN"]
    tomorrow=datetime.now()+timedelta(1)


    suffix=["st","nd","rd","th"]
    date_end_digit=int(tomorrow.strftime("%d"))
    if date_end_digit==1:
        suffix_for_date=suffix[0]
    elif date_end_digit==2:
        suffix_for_date=suffix[1]
    elif date_end_digit==3:
        suffix_for_date=suffix[2]
    else:
        suffix_for_date=suffix[3]
    for_date=tomorrow.strftime("%d{}_%b'%y").format(suffix_for_date)

    print(len(Email_Id))
    
    list_of_dfs=[df2,df,df3]

    for i in list_of_interdomains:
        subject=f"ONLY FOR TESTING: KPI Monitoring | {i} for MPBN CRs | {for_date}"
        if i=="CS Core":
            to=Email_Id.iloc[4]['To Mail List']
            cc=Email_Id.iloc[4]['Copy Mail List']
            dataframe=df2
        elif i=="PS Core":
            to=Email_Id.iloc[3]['To Mail List']
            cc=Email_Id.iloc[3]['Copy Mail List']
            dataframe=df
        elif i=="RAN":
            to=Email_Id.iloc[5]['To Mail List']
            cc=Email_Id.iloc[5]['Copy Mail List']
            dataframe=df3
        mpbn_html_body="""
            <html>
                <body>
                    <div>
                            <p><br><br>Hi Team,</p><br><br>
                            <p>Please find below the list of MPBN activity which includes Core nodes, so KPI monitoring required. Impacted nodes with KPI details given below. Please share KPI monitoring resource from your end.<br><br></p>
                            <p>@Core Team: Please contact below spoc region wise if any issue with KPI input.<br><br></p>
                            <p>Manoj Kumar: North region and west region </p>
                            <p>Arka Maiti: East region and South region <br></p>
                            <p>Note:-If there is any deviation in KPI please call to Executer before 6 AM. After that please call to technical validator/Team Lead.<br><br></p>
                    
                    </div>
                    <div>
                        {}
                    </div>
                    <div>
                            <p>With Regards</p>
                            <p>{}</p>
                            <p>Ericsson India Global Services Pvt. Ltd.</p>
                    </div>
                </body>
            </html>
        """
        sendmail(dataframe,to,cc,mpbn_html_body,subject,sender)

    """
    # this is one way of writing into the excel sheet but it didn't work

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
    
    writer=pd.ExcelWriter(workbook,engine='xlsxwriter')
    daily_plan_sheet.to_excel(writer,sheet_name='Planning Sheet',index=False)
    df.to_excel(writer,sheet_name=sheetname,index=False)
    df2.to_excel(writer,sheet_name=sheetname2,index=False)
    df3.to_excel(writer,sheet_name=sheetname3,index=False)
    Email_Id.to_excel(writer,sheet_name='Mail Id',index=False)

    writer.save()

#####################################################################
############################# Fetch-details #########################
#####################################################################

def fetch_details(sender):
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

            tomorrow=datetime.now()+timedelta(1)

            if daily_plan_sheet.iloc[j]['Execution Date']==tomorrow.strftime("%d/%m/%Y"): # Adding constraint to check for CRs for next date only

                if daily_plan_sheet.iloc[j]['Circle']==circles[i]:

                    execution_date.append(daily_plan_sheet.iloc[j]['Execution Date'])
                    maintenance_window.append(daily_plan_sheet.iloc[j]['Maintenance Window'])
                    cr_no.append(daily_plan_sheet.iloc[j]['CR NO'])
                    activity_title.append(daily_plan_sheet.iloc[j]['Activity Title'])
                    risk.append(daily_plan_sheet.iloc[j]['Risk'])
                    circle.append(daily_plan_sheet.iloc[j]['Circle'])
                    location.append(daily_plan_sheet.iloc[j]['Location'])

            dictionary_for_insertion={'Execution Date':execution_date, 'Maintenance Window':maintenance_window, 'CR NO':cr_no, 'Activity Title':activity_title, 'Risk':risk,'Location':location,'Circle':circle}
            dataframe=pd.DataFrame(dictionary_for_insertion)
            dataframe.reset_index(drop=True,inplace= True)


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
            
    paco_cscore(workbook,sender)


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
    
    except EmailIDError as error:
        print(error)


   