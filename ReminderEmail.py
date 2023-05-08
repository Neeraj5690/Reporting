from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import pandas as pd
import smtplib

Email_Content="Good Evening !!!  Please don't forget to add report file for this week."
FileLink="abc"


ExcelFileName = "UserData"
loc = (ExcelFileName + '.xlsx')
wb = openpyxl.load_workbook(loc)

Sheetname="Data"
sheet = wb[Sheetname]
User_Name_Sheet={}
User_Name_Email={}
for x in range(2, 200):
    if sheet.cell(x,1).value == None:
        break
    else:
        User_Name_Sheet[sheet.cell(x, 1).value]=sheet.cell(x, 2).value
        User_Name_Email[sheet.cell(x, 1).value]=sheet.cell(x, 3).value

#print(User_Name_Sheet)
UserKeys=list(User_Name_Sheet.keys())
#print(UserKeys)
for user in range(0,len(UserKeys)):
    try:
        html = '''
                    <html>
                        <body>
                            <p>Hi '''+UserKeys[user]+'''</p 
                            <p>''' + Email_Content + '''<br /></p>
                            <p>To know more how to add file to GitHub folder checkout the link given below <br /></p>
                            <p>''' + "https://drive.google.com/drive/folders/" + FileLink + '''<br /><br /></p>
                            <p>Many Thanks <br/>Neeraj</p>
                        </body>
                    </html>
                    '''
        email_from = 'Test Automation Team'
        Email_From="neeraj1wayitsol@gmail.com"
        Email_Subject="Weekly Report Reminder Email"

        y = User_Name_Email[UserKeys[user]]
        SenderEmail = Email_From
        date_str = pd.Timestamp.today().strftime('%m-%d-%Y')
        msg = MIMEMultipart()
        msg['Subject'] = Email_Subject + " " + date_str
        msg['From'] = email_from
        msg['To'] = y

        msg.attach(MIMEText(html, "html"))

        # ------------------------To attach all in e-Mail-----------------------
        email_string = msg.as_string()
        # -----------------------------------------------------------------------

        # ----------------------------SMTP setup--------------------------------
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        RandmStr = "tsiajyfnhywxctwi"
        server.login(SenderEmail, RandmStr)
        #server.sendmail(email_from, y, email_string)
        print("Reminder email sent for "+UserKeys[user])
        server.quit()

    except:
        print("Email not sent for "+UserKeys[user])