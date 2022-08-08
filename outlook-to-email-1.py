import os
os.system("pip install lxml html5lib beautifulsoup4")
import win32com.client as wc
import pandas as pd

outlook=wc.Dispatch("Outlook.Application")
mapi=outlook.GetNamespace("MAPI")


for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)

