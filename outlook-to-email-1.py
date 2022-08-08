import os
os.system("pip install lxml html5lib beautifulsoup4")
import win32com.client as wc
import pandas as pd
import subprocess

# Dispatching the outlook application
outlook=wc.Dispatch("Outlook.Application")
mapi=outlook.GetNamespace("MAPI")

# Shows the accounts in outlook
for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)

inbox=mapi.GetDefaultFolder(6) #6 for inbox

# If you have other folders then you can use the below statement
# inbox = mapi.GetDefaultFolder(6).Folders["your_sub_folder"]

messages=inbox.Items

"""Use Restrict function to filter your email message. For instance, 
we can filter by receiving time in past 24 hours, 
and email sender as “contact@codeforests.com” with subject as “Sample Report”"""

"""
received_dt = datetime.now() - timedelta(days=1)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
messages = messages.Restrict("[SenderEmailAddress] = 'contact@codeforests.com'")
messages = messages.Restrict("[Subject] = 'Sample Report'")
"""

#Let's assume we want to save the email attachment to the below directory
user=subprocess.getoutput("echo %username%")
outputDir = rf"C:\Users\{user}\Attachment"
try:
    for message in list(messages):
        try:
            s = message.sender
            for attachment in message.Attachments:
                attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                print(f"attachment {attachment.FileName} from {s} saved")
        except Exception as e:
            print("error when saving the attachment:" + str(e))
except Exception as e:
		print("error when processing emails messages:" + str(e))