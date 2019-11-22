import openpyxl
import re
import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('test case template.xlsx')
worksheet = workbook.add_worksheet("TestCases")


wb = openpyxl.load_workbook('Test cases_DCD_Updated_03212018.xlsx')
sheet = wb.get_sheet_by_name('Install & Uninstall')

list_steps = []
write_row = 3
col = 7   # Column H
for rowOfCellObjects in sheet['F2':'F23']:
    for cellObj in rowOfCellObjects:
        if ((cellObj.coordinate).startswith("F")):
            print("-------------", cellObj.coordinate, "-------------")
            steps = cellObj.value
            list_steps = steps.splitlines()
            for step in list_steps:
                if not step.strip() == "":
                    print(step)
                    worksheet.write(write_row, col, step)
                    write_row += 1

list_steps = []
write_row = 3
col = 8   # Column I
for rowOfCellObjects in sheet['G2':'G23']:
    for cellObj in rowOfCellObjects:
        if ((cellObj.coordinate).startswith("G")):
            print("-------------", cellObj.coordinate, "-------------")
            steps = cellObj.value
            list_steps = steps.splitlines()
            for step in list_steps:
                if not step.strip() == "":
                    print(step)
                    worksheet.write(write_row, col, step)
                    write_row += 1
