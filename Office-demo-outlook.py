import win32com.client as win32
olapp = win32.Dispatch('Outlook.Application')
olNS = olapp.GetNameSpace('MAPI')
mailItem = olapp.CreateItem(0) #creating mailitem
mailItem.Subject = 'TEST'
mailItem.Body = '''Hello there

regards 
Enjoy Maity
'''
mailItem.to = "ankushtomar433@gmail.com"
mailItem._oleobj_.Invoke(*(64209,0,8,0, olNS.Accounts.Item('enjoy.maity@ericsson.com')))
mailItem.Display()
mailItem.Save()
mailItem.Send()
