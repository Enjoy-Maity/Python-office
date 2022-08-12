class EmptyString (Exception):
    def __init__(self,msg):
        self.msg=msg

class ContainsInteger(Exception):
    def __init__(self,msg):
        self.msg=msg


#####################################################################
#############################    Sendmail   #########################
#####################################################################
def get_col_widths(dataframe):
    # First we find the maximum length of the index column   
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns.values  ]

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

def paco_cscore(sender):
    
    user=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

    workbook=r"C:\Users\{}\Daily\MPBN Daily Planning Sheet.xlsx".format(user)
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
        tomorrow=datetime.now()+timedelta(1)

        if daily_plan_sheet.iloc[i]['Execution Date']==tomorrow.strftime("%d-%m-%Y"):
            if daily_plan_sheet.iloc[i]['Domain kpi']=="PS Core" or daily_plan_sheet.iloc[i]['Domain kpi']=="Paco-circle" or daily_plan_sheet.iloc[i]['Domain kpi']=="paco-circle" or daily_plan_sheet.iloc[i]['Domain kpi']=="Paco" or daily_plan_sheet.iloc[i]['Domain kpi']=="ps core" or daily_plan_sheet.iloc[i]['Domain kpi']=="pS Core" or daily_plan_sheet.iloc[i]['Domain kpi']=="Ps core" or daily_plan_sheet.iloc[i]['Domain kpi']=="ps" or daily_plan_sheet.iloc[i]['Domain kpi']=="PS":
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
    df.fillna("NA")


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
        tomorrow=datetime.now()+timedelta(1)

        if daily_plan_sheet.iloc[i]['Execution Date']==tomorrow.strftime("%d-%m-%Y"):
            if daily_plan_sheet.iloc[i]['Domain kpi']=="CS Core" or daily_plan_sheet.iloc[i]['Domain kpi']=="cs core" or daily_plan_sheet.iloc[i]['Domain kpi']=="CS" or daily_plan_sheet.iloc[i]['Domain kpi']=="cs" or daily_plan_sheet.iloc[i]['Domain kpi']=="cS":
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
    df2.fillna("NA")

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
        tomorrow=datetime.now()+timedelta(1)

        if daily_plan_sheet.iloc[i]['Excution Date']==tomorrow.strftime("%d-%m-%Y"):
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
    df3.fillna("NA")

    df.reset_index(drop=True,inplace=True)
    df2.reset_index(drop=True,inplace=True)
    df3.reset_index(drop=True,inplace=True)



    
    writer=pd.ExcelWriter(workbook,engine='xlsxwriter')

    daily_plan_sheet.to_excel(writer,sheet_name='Planning Sheet',index=False)
    df.to_excel(writer,sheet_name=sheetname,index=False)
    df2.to_excel(writer,sheet_name=sheetname2,index=False)
    df3.to_excel(writer,sheet_name=sheetname3,index=False)
    Email_Id.to_excel(writer,sheet_name='Mail Id',index=False)


    workbook=writer.book
    worksheet1=writer.sheets['Planning Sheet']
    worksheet2=writer.sheets[sheetname]
    worksheet3=writer.sheets[sheetname2]
    worksheet4=writer.sheets[sheetname3]
    worksheet5=writer.sheets['Mail Id']
    header_format=workbook.add_format({'bold':True,'fg_color': '#0033cc','font_color':'#ffffff','border':1})
    format=workbook.add_format({'num_format':'dd/mm/yyyy'})

    for i, width in enumerate(get_col_widths(daily_plan_sheet)):
        worksheet1.set_column(i, i, width)

    for col_num, value in enumerate(daily_plan_sheet.columns.values):
        worksheet1.write(0, col_num, value, header_format)
    
    for j in range(0,len(daily_plan_sheet)):
        temp="B"+str(j+2)
        value=daily_plan_sheet.iloc[j][1]
        worksheet1.write(temp,value,format)
    
    for i in range(0,len(daily_plan_sheet)):
        worksheet1.set_column(1,1,15)
    
    for i, width in enumerate(get_col_widths(df)):
        worksheet2.set_column(i, i, width)
    for col_num, value in enumerate(df.columns.values):
        worksheet2.write(0, col_num, value, header_format)
    
    for i, width in enumerate(get_col_widths(df2)):
        worksheet3.set_column(i, i, width)
    for col_num, value in enumerate(df2.columns.values):
        worksheet3.write(0, col_num, value, header_format)
    
    for i, width in enumerate(get_col_widths(df3)):
        worksheet4.set_column(i, i, width)
    for col_num, value in enumerate(df3.columns.values):
        worksheet4.write(0, col_num, value, header_format)
    
    for i, width in enumerate(get_col_widths(Email_Id)):
        worksheet5.set_column(i, i, width)
    for col_num, value in enumerate(Email_Id.columns.values):
        worksheet5.write(0, col_num, value, header_format)
    
    writer.save()

    c=input("Do you want to snd the mails or not? y/n ")
    if c=='y' or c=='Y':
        list_of_interdomains=["CS Core","PS Core","RAN"]
        tomorrow=datetime.now()+timedelta(1)

        suffix=["st","nd","rd","th"]
        date_end_digit=int(tomorrow.strftime("%d"))%10
        date_digits=int(tomorrow.strftime("%d"))%100
        if date_digits<10 or date_digits>20:
            if date_end_digit==1:
                suffix_for_date=suffix[0]
            elif date_end_digit==2:
                suffix_for_date=suffix[1]
            elif date_end_digit==3:
                suffix_for_date=suffix[2]
            else:
                suffix_for_date=suffix[3]
        else:
            suffix_for_date=suffix[3]
        for_date=tomorrow.strftime("%d{}_%b'%y").format(suffix_for_date)
        
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
                                <p>Hi Team,</p>
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
            print(f"\nMail sent for {i}")


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
            paco_cscore(sender)

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
