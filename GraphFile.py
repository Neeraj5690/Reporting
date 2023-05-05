import matplotlib.pyplot as plt
import openpyxl
import pandas as pd

ExcelFileName = "ReportData"
locx = (ExcelFileName + '.xlsx')
wbx = openpyxl.load_workbook(locx)

Sheetname="GlobalData"
sheetx = wbx[Sheetname]

Project_Name = ""
BarGraph_show = "No"
BarGraph_color = "blue"

for ix in range(1, 200):
    if sheetx.cell(ix, 1).value == None:
        break
    else:
        if sheetx.cell(ix, 1).value == "Project_Name":
            print("Project Name is: "+sheetx.cell(ix, 2).value)
            Project_Name=sheetx.cell(ix, 2).value

        if sheetx.cell(ix, 1).value == "BarGraph_color":
            BarGraph_color=sheetx.cell(ix, 2).value
            print("BarGraph_color is "+BarGraph_color)

        if sheetx.cell(ix, 1).value == "BarGraph_show":
            BarGraph_show=sheetx.cell(ix, 2).value
            print("BarGraph_show is "+BarGraph_show)

if BarGraph_show == "Yes":
    Sheetname="ModulesData"
    sheetx = wbx[Sheetname]
    ModuleName=[]
    for ix1 in range(2, 200):
        if sheetx.cell(ix1, 1).value == None:
            break
        else:
            ModuleName.append(sheetx.cell(ix1, 1).value)
    print("ModuleName "+str(ModuleName))

    Column_Name = []
    for ix2 in range(1, 20):
        if sheetx.cell(1,ix2).value == None:
            break
        else:
            Column_Name.append(sheetx.cell(1,ix2).value)
    print("Column_Name "+str(Column_Name))

    Bugs_Count={}
    Bugs_CountList=[]
    ColumnNumber=Column_Name.index("Bugs_Count")+1
    print("ColumnNumber: "+str(ColumnNumber))
    ifNothing=0
    for ix3 in range(1, len(ModuleName)+1):
        if sheetx.cell(ix3+1, 1).value == ModuleName[ix3-1]:
            if sheetx.cell(ix3+1, ColumnNumber).value==None:
                Bugs_Count[ModuleName[ix3 - 1]] = 0
                Bugs_CountList.append(0)
            else:
                Bugs_Count[ModuleName[ix3-1]]=sheetx.cell(ix3+1, ColumnNumber).value
                Bugs_CountList.append(sheetx.cell(ix3+1, ColumnNumber).value)
    print("Bugs_Count "+str(Bugs_Count))
    print("Bugs_CountList "+str(Bugs_CountList))

    # Creating Module Vs Bugs Count Bar Graph
    data = {'modules': ModuleName,
        'bugs': Bugs_CountList
        }
    df = pd.DataFrame(data)
    colors = [BarGraph_color]
    try:
        plt.bar( df['modules'],df['bugs'], color=colors)
    except:
        print("Color is invalid")
        colors = ['blue']
        plt.bar(df['modules'], df['bugs'], color=colors)

    ax = plt.gca()
    plt.draw()
    ax.set_xticklabels(df['modules'] , rotation = 55,fontsize=8)

    plt.title('Module Vs Bugs Count', fontsize=10)
    plt.xlabel('Modules', fontsize=8)
    plt.ylabel('Bugs', fontsize=8)
    plt.grid(False)
    plt.gcf().set_size_inches(5, 7)
    plt.savefig('ModuleVsBugsCount.jpg', dpi=150)

else:
    print("BarGraph_show is " +BarGraph_show)