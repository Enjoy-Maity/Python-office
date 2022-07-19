import os #module to interact with OS
os.system("pip install pywin32 pillow") # installing important python modules
import win32com.client as win32
from PIL import ImageGrab

workbook_path = os.getcwd()+'C:\Users\emaienj\OneDrive - Ericsson\Book1.xlsx'

excel = win32.Dispatch('Excel.Application')

wb = excel.Workbooks.Open(workbook_path)
sheet = wb.Sheets.Item(1)
sheet = wb.Sheets[0]
sheet = wb.Sheets['Sheet1']
