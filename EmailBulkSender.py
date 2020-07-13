import pandas
import smtplib
import time
import os
import sys
import imghdr
import random
from email.message import EmailMessage

EXCEL_DATA = pandas.read_excel('excelname.xlsx', sheet_name='excelnamesheet') #preferably on the same folder as the this script
emails = EXCEL_DATA['Mail'].values    #change the 'Mail' according to your spreadsheet column (the name of your  column)
EMAIL_ADDRESS = "yourgmail@gmail.com" #it will only work with gmail 
EMAIL_PASSWORD = "yourpassword"
for email in emails:
    msg = EmailMessage()
    msg['Subject'] = 'youEmailsubject'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = email
    print(email)
        text = """
        Your message should be here 
        Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec eget lorem consequat, vehicula sapien sit amet, malesuada eros. Aliquam dapibus lectus at
        euismod suscipit. Curabitur vitae rhoncus eros, eget varius purus. Nam vitae magna condimentum, viverra mi vitae, aliquam lorem. Integer non enim justo. 
        Nunc ut ante fringilla, pharetra ligula in, placerat magna. Aenean iaculis tempor erat vel tristique. Proin sodales est quis lacinia aliquam. Etiam non laoreet ante.
    """
    msg.set_content(text)
    files = ['yourPDFfile.pdf'] # preferably on the same folder as this script
    for file in files:
        with open(file,'rb') as f:
            file_data = f.read()
            file_name = f.name
        msg.add_attachment(file_data,maintype='application', subtype='octet-stream', filename=file_name)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as tmp :
        tmp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
        tmp.send_message(msg)
        rdm = random.randrange(200,10,-2)
        print(rdm)
        time.sleep(rdm)  
