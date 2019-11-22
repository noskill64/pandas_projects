import pandas as pd
# from openpyxl.workbook import workbook


my_sheet = 'Settings'
file_name = '/Users/Admin/Desktop/import_testcases/Reviewed Test cases_DCD_Updated_19062018.xlsx'  # name of your excel file
df = pd.read_excel(file_name, sheet_name=my_sheet)
# print(df)
print(df.head())
