import time
import glob
from pathlib import Path
import os
import openpyxl

home = str(Path.home())
home = home.replace(os.sep, '/')
print(home)

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

print(User_Name_Sheet)
UserKeys=list(User_Name_Sheet.keys())
print(UserKeys)
for user in range(0,len(UserKeys)):
    try:
        os.remove(home+'/.jenkins/workspace/Create_Graph/'+UserKeys[user]+'_ModuleVsBugsCount.jpg')
    except:
        print("ModuleVsBugsCount not able to delete for "+UserKeys[user])
#-----------------To delete pdf and report xlsx files----------------------------
time.sleep(2)
ii=0
fileList1 = glob.glob(home+'/.jenkins/workspace/CreateReport/*.pdf')
fileList2 = glob.glob("ReportData1/"+'/*.xlsx')

for ii in range(0,len(fileList1)):
    try:
        os.remove(fileList1[ii])
        os.remove(fileList2[ii])
    except Exception as ae:
        print(ae)
        print("No Attachment found to delete")

