import glob
import os
import smtplib
import ssl
import time
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import pandas as pd
import ast
import sys
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pandas as pd


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
        print("User_Name : "+sheet.cell(x, 1).value)
        print("User_File : "+sheet.cell(x, 2).value)
        print("User_Email : " + sheet.cell(x, 3).value)
        User_Name_Sheet[sheet.cell(x, 1).value]=sheet.cell(x, 2).value
        User_Name_Email[sheet.cell(x, 1).value]=sheet.cell(x, 3).value

print(User_Name_Sheet)
UserKeys=list(User_Name_Sheet.keys())
print(UserKeys)
for user in range(0,len(UserKeys)):
    # print(UserKeys[user])
    # print(User_Name[UserKeys[user]])

    try:
        #-------------------To read content to send in e-Mail--------------------
        # Connecting with Main Report Data File
        ExcelFileName = "ReportData1/"+User_Name_Sheet[UserKeys[user]]
        locx = (ExcelFileName + '.xlsx')
        wbx = openpyxl.load_workbook(locx)

        # Reading GlobalData tab of Main Report Data File
        date_str = pd.Timestamp.today().strftime('%m-%d-%Y')
        Sheetname = "GlobalData"
        sheetx = wbx[Sheetname]
        for ix in range(1, 200):
            if sheetx.cell(ix, 1).value == None:
                break
            else:
                if sheetx.cell(ix, 1).value == "Report_Name":
                    print("Report_Name is: " + sheetx.cell(ix, 2).value)
                    FileDelete=sheetx.cell(ix, 2).value
                    Report_Name = sheetx.cell(ix, 2).value+'_'+date_str+'.pdf'
                if sheetx.cell(ix, 1).value == "Email_From":
                    print("Email_From is: " + sheetx.cell(ix, 2).value)
                    Email_From = sheetx.cell(ix, 2).value
                if sheetx.cell(ix, 1).value == "Email_To":
                    print("Email_To is: " + sheetx.cell(ix, 2).value)
                    Email_To = sheetx.cell(ix, 2).value
                if sheetx.cell(ix, 1).value == "Email_Subject":
                    print("Email_Subject is: " + sheetx.cell(ix, 2).value)
                    Email_Subject = sheetx.cell(ix, 2).value
                if sheetx.cell(ix, 1).value == "Email_Content":
                    print("Email_Content is: " + sheetx.cell(ix, 2).value)
                    Email_Content = sheetx.cell(ix, 2).value
                if sheetx.cell(ix, 1).value == "GoogleAppCode":
                    print("GoogleAppCode is: " + sheetx.cell(ix, 2).value)
                    GoogleAppCode = sheetx.cell(ix, 2).value
                if sheetx.cell(ix, 1).value == "GoogleDriveFolderID":
                    print("GoogleDriveFolderID is: " + sheetx.cell(ix, 2).value)
                    GoogleDriveFolderID = sheetx.cell(ix, 2).value

        html = '''
            <html>
                <body>
                    <p>Hi Team</p 
                    <p>'''+Email_Content+'''<br /></p>
                    <p>To access old reports go to drive link given below <br /></p>
                    <p>'''+"https://drive.google.com/drive/folders/"+GoogleDriveFolderID+'''<br /><br /></p>
                    <p>Many Thanks <br/>'''+UserKeys[user]+'''</p>
                </body>
            </html>
            '''

        def attach_file_to_email(msg, attach,Report_Name, extra_headers=None):
            with open(attach, "rb") as f:
                file_attachment = MIMEApplication(f.read())
            file_attachment.add_header(
                "Content-Disposition",
                f"attachment; filename= {Report_Name}",
            )
            if extra_headers is not None:
                for name, value in extra_headers.items():
                    file_attachment.add_header(name, value)
            msg.attach(file_attachment)

        email_from = 'Test Automation Team'
        y = Email_To
        email_to = ast.literal_eval(y)
        SenderEmail=Email_From
        date_str = pd.Timestamp.today().strftime('%m-%d-%Y')
        msg = MIMEMultipart()
        msg['Subject']=Email_Subject+" "+date_str
        msg['From'] = email_from
        msg['To'] = ','.join(email_to)
        msg.attach(MIMEText(html, "html"))
        attach_file_to_email(msg, Report_Name,Report_Name)

        #-----------------------------------------------------------------------

        # ------------------------To attach all in e-Mail-----------------------
        email_string = msg.as_string()
        context = ssl.create_default_context()
        # -----------------------------------------------------------------------

        # ----------------------------SMTP setup--------------------------------
        server=smtplib.SMTP_SSL('smtp.gmail.com',465)
        RandmStr=GoogleAppCode
        server.login(SenderEmail,RandmStr)
        #server.sendmail(email_from, email_to, email_string)
        print("Test Report sent")
        server.quit()
        #--------------------------GDrive setup-----------------------------------
        gauth = GoogleAuth()
        gauth.LoadCredentialsFile("mycreds.txt")
        if gauth.credentials is None:
            gauth.LocalWebserverAuth()
        elif gauth.access_token_expired:
            gauth.Refresh()
        else:
            gauth.Authorize()
        gauth.SaveCredentialsFile("mycreds.txt")
        drive = GoogleDrive(gauth)
        #--------------------------GDrive upload-----------------------------------
        upload_file_list = [Report_Name]
        for upload_file in upload_file_list:
            gfile = drive.CreateFile({'parents': [{'id': GoogleDriveFolderID}]})
            gfile.SetContentFile(upload_file)
            gfile.Upload()
            #--------------------------------------------------------------------------

        #-----------------To delete pdf and report files----------------------------
        time.sleep(2)
        ii=0
        fileList = glob.glob('*.pdf')
        for ii in range(0,len(fileList)):
            try:
                os.remove(fileList[ii])
            except Exception as ae:
                print(ae)
                print("No Attachment found to delete")
        os.remove("ModuleVsBugsCount.jpg")
        #-----------------------------------------------------------------------

        sheetx.cell(row=1, column=5).value = date_str
        wbx.save(locx)

    except Exception as aaa:
        print("Report File not found for "+UserKeys[user])
        print(aaa)

        Email_Content="Good Evening !!!  Please don't forget to add report file for this week."
        FileLink="abc"

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
        # def attach_file_to_email(msg, attach, Report_Name, extra_headers=None):
        #     with open(attach, "rb") as f:
        #         file_attachment = MIMEApplication(f.read())
        #     file_attachment.add_header(
        #         "Content-Disposition",
        #         f"attachment; filename= {Report_Name}",
        #     )
        #     if extra_headers is not None:
        #         for name, value in extra_headers.items():
        #             file_attachment.add_header(name, value)
        #     msg.attach(file_attachment)


        email_from = 'Test Automation Team'
        y = User_Name_Email[UserKeys[user]]
        email_to = ast.literal_eval(y)
        SenderEmail = Email_From
        date_str = pd.Timestamp.today().strftime('%m-%d-%Y')
        msg = MIMEMultipart()
        msg['Subject'] = Email_Subject + " " + date_str
        msg['From'] = email_from
        msg['To'] = ','.join(email_to)
        msg.attach(MIMEText(html, "html"))

        # -----------------------------------------------------------------------

        # # ------------------------To attach all in e-Mail-----------------------
        # email_string = msg.as_string()
        # context = ssl.create_default_context()
        # -----------------------------------------------------------------------

        # ----------------------------SMTP setup--------------------------------
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        RandmStr = "tsiajyfnhywxctwi"
        server.login(SenderEmail, RandmStr)
        # server.sendmail(email_from, email_to, email_string)
        print("Test Report sent")
        server.quit()