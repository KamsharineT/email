##########################################################################################################
##Developer : Kamsharine Thayananthan                                                                   ##
##Purpose   : This is written to send a consildate mail of all alarms to users                 			##
##Date      : 2023/04/26                                                                                ##
##########################################################################################################

import psycopg2
import xlwt
from xlwt import Workbook 
from datetime import datetime,timedelta
import os
import smtplib
from prettytable import PrettyTable
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import pandas as pd
import time

# outlook=win32.Dispatch('outlook.application')
# mail=outlook.CreateItem(0)

# mail.To='kamsharine@gmail.com'
# mail.Subject='Message subject'
# mail.Body='Message body'
# mail.HTMLBody='<h2>Alarm List of Today </h2>' #this field is optional


# # To attach a file to the email (optional):
# attachment="Path to the attachment"
# mail.Attachments.Add(attachment)
  # print(datatext)




################################################################
#This function is to read the recipients from excel sheet     ##
################################################################
def recipients():
  df = pd.read_excel(io='recipients.xlsx',dtype={'Email':object})
  print(df)
  return df


################################################################
#This function is to write the user flagged to a table        ##
################################################################
def writetotable():
  cursor.execute("SELECT b.tag, b.name,a.created,a.reported_by_id, a.scheduled_maintennce,a.action,a.id FROM alarms_alarm a JOIN devices_device b ON a.device_id = b.id where a.reported_by_id is not null and deleted is null and end_date is null order by a.device_id")
  result = cursor.fetchall()
  print(result)
  alarmtable = PrettyTable(["No","Device tag","Station","Date Reported","Deadline","Reported By","Fault","Scheduled Maintenance","Action"])

  if result == []:
    alarmtable.add_row(["1",'N/A','N/A','N/A','N/A','N/A','N/A','N/A','N/A'])

  for j in range(len(result)):
    
    reportedbyid = result[j][3]

    cursor.execute("select first_name||' '||last_name from users_user where id = '"+str(reportedbyid)+"'")
    reportedby = cursor.fetchall()
    reportedby = reportedby[0][0]

    alarmid = result[j][6]
    cursor.execute("select fault_message from alarms_alarm_detail where alarm_id = '"+str(alarmid)+"'")
    fault = cursor.fetchall()
    fault = fault[0][0]

    try:
      scheduled = datetime.strftime(result[j][4]+timedelta(hours=(+8)),'%Y-%m-%d %H:%M:%S.%f+00:00')
    except:
      scheduled = "Not Defined"

    action = result[j][5]
    # print("action ",action)
    if (action == "" or action is None):
      print("IF")
      action = "N/A"
    else:
      action = action
    
    repdate = (result[j][2])+timedelta(hours=(+8))
    alarmtable.add_row([j+1,result[j][0],result[j][1],datetime.strftime(repdate,'%Y-%m-%d %H:%M:%S.%f+00:00'),datetime.strftime((repdate+timedelta(days=(+3))),'%Y-%m-%d %H:%M:%S.%f+00:00'),reportedby,fault,str(scheduled),action])

    print("result[j]")
    print(result[j][0])

    if result[j][0] is None:
      alarmtable.add_row([j,'N/A','N/A','N/A','N/A','N/A','N/A','N/A','N/A'])
  # print(alarmtable.get_html_string())
  sendmail(alarmtable)

################################################################
#This function is to write the user flagged alarms to an excel##
################################################################
def writetoexcel():
  cursor.execute("SELECT b.tag,b.name,a.created,a.reported_by_id, a.scheduled_maintennce,a.action,a.id FROM  alarms_alarm a JOIN devices_device b ON a.device_id = b.id and deleted is null and end_date is null order by a.device_id")
  result = cursor.fetchall()
  wb = Workbook()

  sheet = wb.add_sheet('List of station alarms')
  # print(type(result[1]))
  sheet.write(0,0,"No")
  sheet.write(0,1,"Device Tag")
  sheet.write(0,2,"Station")
  sheet.write(0,3,"Date Reported")
  sheet.write(0,4,"Deadline")
  sheet.write(0,5,"Reported By")
  sheet.write(0,6,"Fault")
  sheet.write(0,7,"Scheduled Maintenance")
  sheet.write(0,8,"Action")

  for i in range(len(result)):  
    sheet.write(i+1,0,i+1)
    sheet.write(i+1,1,result[i][0])
    sheet.write(i+1,2,result[i][1])
    sheet.write(i+1,3,datetime.strftime(((result[i][2])+timedelta(hours=(+8))),'%Y-%m-%d %H:%M:%S.%f+00:00'))
    deadline = (result[i][2])+timedelta(hours=(+8))+timedelta(days=(+3))
    sheet.write(i+1,4,datetime.strftime(deadline,'%Y-%m-%d %H:%M:%S.%f+00:00'))

    alarmid = result[i][6]
    
    
    reportedbyid = result[i][3]
    print(str(reportedbyid))
    if(reportedbyid is not None):
      cursor.execute("select first_name||' '||last_name from users_user where id = '"+str(reportedbyid)+"'")
      reportedby = cursor.fetchall()                                                                                                                               
      reportedby = reportedby[0][0]
    else:
      reportedby = "System"

    sheet.write(i+1,5,reportedby)

    cursor.execute("select fault_message from alarms_alarm_detail where alarm_id = '"+str(alarmid)+"'")
    fault = cursor.fetchall()
    try:
      fault = fault[0][0]
    except:
      fault = "N/A"
    sheet.write(i+1,6,fault)
    try:
      sheet.write(i+1,7,datetime.strftime(result[i][4],'%Y-%m-%d %H:%M:%S.%f+00:00'))
    except:
      sheet.write(i+1,7,"Not Defined")

    action = result[i][5]
    print("action ",action)
    if (action == "" or action is None):
      print("IF")
      sheet.write(i+1,8,"N/A")
    else:
      print("else")
      sheet.write(i+1,8,action)

  wb.save('excel/List_of_Alarms.xls')


##########################################################################
#  This function is to write send the excel and tables to users          #
##########################################################################
def sendmail(alarmtable):

  writetoexcel()
  alarms = alarmtable.get_html_string()

  html = """\
    <html>
        <head>
        <style>
            table, th, td {
                border: 1px solid black;
                border-collapse: collapse;
            }
            th, td {
                padding: 7px;
                text-align: center;    
            }    
        </style>
        </head>
    <body>
    
      <h3> Hi,<br><br>
      Please refer to the attachment for alarm details and the details of user flagged alarms are in the table. <br></h3>
       <br>
       %s
       <br><br>
       <h3>
       Kamsharine</h3>
    </p>
    </body>
    </html>
    """ % (alarms)
  
  print(alarmtable)
  # creates SMTP session
  

  df = recipients()
  # print(len(recipients))

  for i in df.index:
    sendto = (df['Email'][i])
    # cc = ['kamsharine@gmail.com','']
    s = smtplib.SMTP('smtp-mail.outlook.com', 587)
  # start TLS for security
    s.starttls() 

    s.login("kamsharine@outlook.com", "")
  
    sendfrom = "kamsharine@outlook"
    # sendto = "kamsharine@gmail.com"
    
    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')

    todayutc= datetime.strftime(datetime.today(),"%y-%m-%d")

    msg['Subject'] = "Daily Water Quality Profiling Station Overview"
    msg['From'] = sendfrom
    msg['To'] = sendto

    part2 = MIMEText(html,"html")

    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    msg.attach(part2)

    with open("excel/List_of_Alarms.xls", "rb") as f:
      file_data = f.read()
    
    # Add the file as an attachme+nt to the message object
    attachment = MIMEApplication(file_data, _subtype="xls")
    attachment.add_header("content-disposition", "attachment", filename="List_of_Alarms.xls")
    msg.attach(attachment)
    
    # Send the message via local SMTP server.
    # s = smtplib.SMTP('localhost')
    # sendmail function takes 3 arguments: sender's address, recipient's address

    print("sending mail")
    # and message to send - here it is sent as one string.
    s.sendmail(sendfrom, sendto, msg.as_string()) # commented for testing
    time.sleep(30)
  s.quit()
  # os.chdir('excel')


#############################################################################
#                             Main Call                                     #
#############################################################################

while True:
  conn = psycopg2.connect(
    host="",
    database="",
    port="",
    user="",
    password="")
  
  cursor = conn.cursor()
  # today= datetime.strptime("2023-04-26 01:29 PM","%Y-%m-%d %H:%M %p") #testing

  # print(datetime.utcnow()+timedelta(hours=(+8)))
  today = (datetime.utcnow()+timedelta(hours=(+8)))

  # print(today.hour,today.minute)
  # if(1==1):
  if(today.hour == 10 and today.minute==00):
    writetotable()
    conn.close()
  # time.sleep(30)
 
   