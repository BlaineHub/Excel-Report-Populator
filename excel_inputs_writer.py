from email.mime.application import MIMEApplication
import pandas as pd
import datetime
import shutil
from email.mime.text import MIMEText
import smtplib
from os.path import basename

#Step1 - Loading Excel template & Creating New Template
Current_Date = datetime.datetime.today().strftime ('%d-%b-%Y')
original = 'template.xlsx'
target = 'template_'+ str(Current_Date) + '.xlsx'
shutil.copyfile(original,target)
print('EXCEL FILE CREATED FROM THE TEMPLATE')


#Step2 - Loading DataFile (COULD BE DIRECTLY ATTACHED TO DB QUERY)
df = pd.read_csv('data.csv', low_memory=False, encoding='utf-8')
print('DATA LOADED')

#Step3 - spliting data to seperate dateframes to write to sheets
virtual_meters = df.loc[df.FileInput.isin(['0'])]
devices = list(df['FileInput'].unique())
r_devices = [dev for dev in devices if dev != '0' and not dev.startswith('E')]
w_devices = [dev for dev in devices if dev.startswith('E')]
real_meters = df.loc[df.FileInput.isin(r_devices)]
weather = df.loc[df.FileInput.isin(w_devices)]
print('DATA SPLIT INTO VARIABLES PER SHEET')

#Step4 - WRITING DATA TO EXCEL TEMPLATE & SAVING CONTENTS
writer = pd.ExcelWriter('template_'+ str(Current_Date) + '.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
real_meters.to_excel(writer, sheet_name='sheet1',index=False)
virtual_meters.to_excel(writer, sheet_name='sheet2',index=False)
weather.to_excel(writer, sheet_name='sheet3',index=False)
writer.save()
print('FILES WRITTEN TO EXCEL')


#Step5 - Email service
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication

from_email='my email'
from_password=''
to_email= 'to email'

mail_content = '''Please see attached Report,
Kind Regards,
'''

message = MIMEMultipart()
attach_file_name = 'template_'+ str(Current_Date) + '.xlsx'
message['To']=to_email
message['From']=from_email
message['Subject'] = basename(attach_file_name)
message.attach(MIMEText(mail_content))


attach_file = open(attach_file_name, 'rb')
payload = MIMEBase('application', 'octet-stream' ,name= basename(attach_file_name))
payload.set_payload((attach_file).read())
encoders.encode_base64(payload) 
message.attach(payload)


gmail=smtplib.SMTP('smtp.gmail.com',587)
gmail.ehlo()
gmail.starttls()
gmail.login(from_email,from_password)
gmail.send_message(message)
gmail.close()
print('MAIL SENT')