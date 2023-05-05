import os
import openpyxl
from fpdf import FPDF, fpdf
import datetime
import pandas as pd
from pathlib import Path

home = str(Path.home())
home = home.replace(os.sep, '/')
print(home)

version = " version: 1.0 "


ExcelFileName = "UserData"
loc = (ExcelFileName + '.xlsx')
wb = openpyxl.load_workbook(loc)

Sheetname="Data"
sheet = wb[Sheetname]
User_Name={}
for x in range(2, 200):
    if sheet.cell(x,1).value == None:
        break
    else:
        print("User_Name : "+sheet.cell(x, 1).value)
        print("User_File : "+sheet.cell(x, 2).value)
        User_Name[sheet.cell(x, 1).value]=sheet.cell(x, 2).value

print(User_Name)
UserKeys=list(User_Name.keys())
print(UserKeys)
for user in range(0,len(UserKeys)):
    # print(UserKeys[user])
    # print(User_Name[UserKeys[user]])

    try:
        ExcelFileName = "ReportData/"+User_Name[UserKeys[user]]

        Report_Title="Report"
        Project_Name = "QA"
        Report_Name = "Report"

        date_str = pd.Timestamp.today().strftime('%m-%d-%Y')
        ct = datetime.datetime.now().strftime("%d_%B_%Y_%I_%M%p")
        ctReportHeader = datetime.datetime.now().strftime("%d %B %Y %I %M %p")

        # Connecting with Main Report Data File
        locx = (ExcelFileName + '.xlsx')
        wbx = openpyxl.load_workbook(locx)

        # Reading GlobalData tab of Main Report Data File
        Sheetname="GlobalData"
        sheetx = wbx[Sheetname]
        for ix in range(1, 200):
            if sheetx.cell(ix, 1).value == None:
                break
            else:
                if sheetx.cell(ix, 1).value == "Project_Name":
                    print("Project Name is: "+sheetx.cell(ix, 2).value)
                    Project_Name=sheetx.cell(ix, 2).value
                if sheetx.cell(ix, 1).value == "Report_Title":
                    print("Report Title is: "+sheetx.cell(ix, 2).value)
                    Report_Title=sheetx.cell(ix, 2).value
                if sheetx.cell(ix, 1).value == "Report_Name":
                    print("Report_Name is: "+sheetx.cell(ix, 2).value)
                    Report_Name=sheetx.cell(ix, 2).value+'_'+date_str+'.pdf'

        # Reading Column Name data
        Sheetname="ModulesData"
        sheetx = wbx[Sheetname]
        Column_Name = []
        for ix2 in range(1, 20):

            if sheetx.cell(1,ix2).value == None:
                break
            else:
                Column_Name.append(sheetx.cell(1,ix2).value)
        print("Column_Name "+str(Column_Name))

        # Reading Module Name data
        ModuleName=[]

        # Reading Bugs_Count data
        Bugs_Count={}
        Bugs_CountList=[]

        # Reading Bugs_Links data
        Bugs_Links={}

        # Reading Triage data
        Comment={}

        for ix1 in range(1, len(Column_Name)+1):
            if ix1==1:
                LastCell=100
            else:
                LastCell = len(ModuleName)

            for ix11 in range(2,LastCell+2 ):
                # Do not change this - it will fetch all the module names
                if ix1==1:
                    if sheetx.cell(ix11, ix1).value == None:
                        break
                    else:
                        ModuleName.append(sheetx.cell(ix11,ix1).value)
                else:
                    # Add new parameter here if required
                    if ix1==2:
                        if ix11 == len(ModuleName)+2:
                            break
                        if sheetx.cell(ix11,ix1).value == None:
                            Bugs_Count[ModuleName[ix11 - 2]] = "0"
                            Bugs_CountList.append("0")
                        else:
                            Bugs_Count[ModuleName[ix11 - 2]] = sheetx.cell(ix11,ix1).value
                            Bugs_CountList.append(sheetx.cell(ix11,ix1).value)
                    if ix1==3:
                        if ix11 == len(ModuleName)+2:
                            break
                        if sheetx.cell(ix11,ix1).value == None:
                            Bugs_Links[ModuleName[ix11 - 2]] = "None"
                        else:
                            Bugs_Links[ModuleName[ix11 - 2]] = sheetx.cell(ix11,ix1).value
                    if ix1==4:
                        if ix11 == len(ModuleName)+2:
                            break
                        if sheetx.cell(ix11,ix1).value == None:
                            Comment[ModuleName[ix11 - 2]] = "None"
                        else:
                            Comment[ModuleName[ix11 - 2]] = sheetx.cell(ix11,ix1).value

        print("Modules are "+str(ModuleName))
        print("Bugs_Count "+str(Bugs_Count))
        MaxBugs = max(Bugs_Count, key=Bugs_Count.get)
        print("MaxBugs "+str(MaxBugs))
        print("Bugs_CountList "+str(Bugs_CountList))
        print("Sum of Bugs_CountList "+str(sum(Bugs_CountList)))
        print("Bugs_Links are "+str(Bugs_Links))
        print("Comments are "+str(Comment))

        class PDF(FPDF):
            def header(self):
                self.image('Logo.png', 10, 8, 33)
                self.set_font('Arial', 'B', 15)
                w = self.get_string_width(Report_Title+": "+Project_Name) + 6
                self.set_x((210 - w) / 2)
                self.set_draw_color(0, 80, 180)
                self.set_fill_color(0,76,153)
                self.set_text_color(255,255,255)
                self.set_line_width(1)
                self.cell(w, 9, Report_Title+": "+Project_Name, 1, 1, 'C', 1)
                self.ln(10)

            def footer(self):
                self.set_y(-15)
                self.set_font('Arial', 'I', 8)
                self.set_text_color(128)
                self.cell(0, 10, 'Page ' + str(self.page_no()), 0, 0, 'C')

            def section_counts(self):
                self.ln(17)
                self.set_font("Arial", size=12)
                w1 = self.get_string_width(" Total Bugs : "+str(sum(Bugs_CountList))) + 3
                X = (70 - w1) / 2
                self.set_x(X)

                Y = 36
                multplyVar=15
                self.set_y(Y)
                self.set_draw_color(0, 80, 180)
                self.set_fill_color(198, 224, 228)
                self.set_text_color(0, 0, 0)
                self.set_line_width(1)
                self.cell(w1, 9, " Total bugs : "+str(sum(Bugs_CountList)), 1, 1, 'L', 1)

                Y=Y+multplyVar
                self.set_y(Y)
                self.set_draw_color(0, 80, 180)
                self.set_fill_color(231, 203, 149)
                self.set_text_color(0, 0, 0)
                self.set_line_width(1)
                w2 = self.get_string_width(" Max bugs from : "+MaxBugs) + 3
                self.cell(w2, 9, " Max bugs from : "+MaxBugs, 1, 1, 'L', 1)

                if Y<=112:
                    Y=112
                else:
                    Y = Y + multplyVar+15
                self.set_y(Y)
                #self.ln(49)
                self.set_font('Arial', 'B', 12)
                self.set_text_color(0, 0, 0)
                self.set_fill_color(200, 220, 255)
                self.cell(0, 0, 'Module details : ', 0, fill=False)
                self.ln(5)

            def section_title(self, num, label):
                self.set_font('Arial', '', 12)
                self.set_text_color(0, 0, 0)
                self.set_fill_color(200, 220, 255)
                self.cell(0, 10, 'Module %d : %s' % (num, label), 0, 1, 'L', 1)
                self.cell(0, 10, 'Bugs Count : '+str(Bugs_Count[label]), 0, 1)
                self.set_text_color(0,0,255)
                self.cell(0, 6, 'Link : ' + str(Bugs_Links[label]), 0, 1)
                self.set_text_color(0, 0, 0)
                self.multi_cell(0, 6, 'Comment : ' + str(Comment[label]), 0)
                self.ln(5)

            def Graph(self):
                self.set_font("Arial", size=8)
                self.set_text_color(80, 80, 80)
                self.cell(0, 1, version +"    Report Date: "+ctReportHeader, 0, 0, 'L')
                try:
                    self.image(home+'.jenkins/workspace/Create_Graph/Neeraj_ModuleVsBugsCount.jpg', 100, 25, 100,90)
                except Exception as aa:
                    print(aa)
                    self.image('img.jpg', 100, 25, 100, 90)

            def print_Data(self, num, Report_Title):
                self.section_title(num, Report_Title)

        pdf = PDF()
        pdf.set_title(Report_Title)
        pdf.add_page()
        pdf.Graph()
        pdf.section_counts()

        # Module Data
        Sheetname="ModulesData"
        sheetx = wbx[Sheetname]
        for ix in range(1, len(ModuleName)+1):
            pdf.print_Data(ix, ModuleName[ix-1])

        pdf.output(Report_Name)

    except Exception as aaa:
        print("Report File not found for "+UserKeys[user])
        print(aaa)