import os #module to interact with OS
os.system("pip install pywin32 pillow") # installing important python modules
import win32com.client as win32
from PIL import ImageGrab

workbook_path = "c:\\users\\emaienj\\downloads\\Book1.xlsx"

excel = win32.Dispatch('Excel.Application')

wb = excel.Workbooks.Open(workbook_path)
sheet = wb.Sheets.Item(1)
sheet = wb.Sheets[0]
sheet = wb.Sheets['Sheet1']
excel.visible= 1
copyrange=sheet.Range('A2:F2')
sheet=wb.Worksheets(1)
sender=sheet.Range("E2").Value
circle = sheet.Range("D2").Value
recipient = sheet.Range("F2").Value
copyrange.CopyPicture(Appearance=1, Format=2)
ImageGrab.grabclipboard().save('paste.png')

excel.Quit()


#Creating Outlook Mail and inserting the Excel content
image_path= os.getcwd()+'\\paste.png'
html_body="""
<div><p>Hi team,</p><br><br>

<p>Please confirm below points so that we will approve CR’s.<br><br></p>

<p>1)  End nodes and service details are required which are running on respective MPBN device (in case of changes on Core/PACO/HLR devices ).<br></p>
<p>2)  Design Maker & Checker confirmation mail need to be shared for all planned activity on Core/PACO/HLR devices.<br></p>
<p>3)  KPI & Tester details need to be shared for all impacted nodes in Level-1 CR’s (SA).Also same details need to be shared for all Level-2 CR’s (NSA) with respect to changes on Core/PACO/HLR devices.<br><br></p>
</div>
<div>
    <img src={}><br><br></img>
</div>
<div>
    <p>Regards<br></p>
    <p>{}</p>
</div>
"""

outlook=win32.Dispatch('Outlook.Application')
msg= outlook.CreateItem(0)
msg.To=recipient
msg.Subject = "Connected End Nodes and their services on MPBN devices: {}".format(circle)
msg.HTMLBody = html_body.format(image_path,sender)
msg.Display()
msg.Save()

msg.Send()