import openpyxl as opxl
import json

data = []

def execPgm():
    try:
        wb = opxl.load_workbook('sample.xlsx')
    except: 
        print("Didn't find the indicated file, create new one")
        # create a new one
        wb = opxl.Workbook()

    # print all the sheet names
    print("All sheets: {}".format(wb.sheetnames))

    # Get the active sheet
    # also could use wb['sheetname'] to get the active tab
    actSheet = wb.active
    print(actSheet.title)

    #change sheet name and color
    actSheet.title = "First sheet"
    actSheet.sheet_properties.tabColor = "1072BA"

    #default sheet title define and append to excel
    defaultSheetTitles = ["title1", "title2", "title3"]

    #Insert titles at row1
    for index, x in enumerate(defaultSheetTitles):
        actSheet.cell(row = 1, column = (index+1), value = x)
    
    
    wb.save('sample.xlsx')

if __name__ == "__main__":
    execPgm()