'''
Authors: Zineb ALAHEM & Monsef ALAHEM
sumerize: the progam send a message with attachements to all emails found in excel
'''

#excel tools
import xlrd
from xlrd import open_workbook

#auto email tools
from email.mime.text import MIMEText
#from email.header import Header
from smtplib import SMTP_SSL

# Import smtplib for the actual sending function
import smtplib

# And imghdr to find the types of our images
import imghdr

# Here are the email package modules we'll need
from email.message import EmailMessage

import time


# email sending server
host_server = 'smtp.gmail.com' #another exmeple depending on your email provider 'smtp.exmail.qq.com'
sender_mail = 'your_mail@gmail.com'
sender_passcode = 'your_pass'


def send_mail(receiver='', mail_title='', mail_content='',files=None):
    # ssl login
    smtp = SMTP_SSL(host_server, 465) #465 587
    print("ssl session success !")
    # set_debuglevel() for debug, 1 enable debug, 0 for disable
    # smtp.set_debuglevel(1)
    smtp.ehlo(host_server)
    print("ehlo success !")
    smtp.login(sender_mail, sender_passcode)
    print("login success !")

    # construct message
    msg = EmailMessage()
    msg.set_content(mail_content)
    msg["Subject"] = mail_title
    msg["From"] = sender_mail
    msg["To"] = receiver

    for file in files:
        with open(file, 'rb') as fp:
            data = fp.read()
        msg.add_attachment(data,maintype='Excel', subtype='document',filename=file)

    smtp.sendmail(sender_mail, receiver, msg.as_string())
    smtp.quit()
    
#loading excel file
wb = open_workbook ("your_contacts.xlsx")
sheet = wb.sheet_by_index(0)

#check numbers of rows
print(sheet.nrows)

# loop over the excel file's rows
for rownum in range(sheet.nrows -2):

    #retriev data from the first collum, in this case the name
    name = sheet.cell(rownum+1,11).value

    #retriev the data from the second collum, in this case the email
    receiver_email = sheet.cell(rownum+1,12).value

    # receiver mail
    receiver = receiver_email
    # mail contents
    mail_content = "Hello,\n\nYour name is " + str(name)
    # mail title
    mail_title = 'say hello'
    print(name)
    send_mail(receiver=receiver\
        ,mail_title=mail_title,\
        mail_content="hello ! im'auto-mailler",\
         files={'your_file1.pdf',\
         'your_file2.docx'})
