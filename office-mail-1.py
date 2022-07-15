import os
import yagmail
#import smtplib
#from email.mime.application import MIMEApplication
#from email.mime.multipart import MIMEMultipart
#from email.mime.base import MIMEBase

def send_email(User,Password,SEND_FROM, SEND_TO, SUBJECT, MAIL_BODY,app_password):
    #multipart=MIMEMultipart()
    #multipart['From']=SEND_FROM
    #multipart['To']=SEND_TO
    #multipart['Subject']=SUBJECT
    #multipart['Body']=MAIL_BODY

    #mailserver=smtplib.SMTP(host='smtp.google.com', port=587)
    #mailserver.ehlo()
    #mailserver.starttls()
    #mailserver.ehlo()
    #mailserver.login(User,Password)

    #mailserver.sendmail(SEND_FROM, SEND_TO, multipart.as_string())
    #mailserver.quit()
    with yagmail.SMTP(User, app_password) as yag:
        yag.send(SEND_TO, SUBJECT, MAIL_BODY)
        print('Sent email successfully')

user="enjoymaity@gmail.com"
password="heqt65IR##heqt65IR$$"
SEND_FROM="enjoymaity@gmail.com"
SEND_TO="ankushtomar433@gmail.com"
#SEND_TO="karan.k.loomba@ericsson.com"
SUBJECT="PYTHON AUTOMATION TEST"
MAIL_BODY="""Hello from Enjoy testing AUTOMATION

Regards
Enjoy Maity"""
app_password= "zieabvkdyutdgwog"
send_email(user, password, SEND_FROM, SEND_TO, SUBJECT, MAIL_BODY, app_password)