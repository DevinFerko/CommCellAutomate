# This python program will help to automate some of the steps required to perform the weekly comm's cell report

# Imports Libraries/Dependencies
import openpyxl

# Path of excel sheet
path = r"C:\Users\U360269\Desktop\Codes\CommCellAutomate\SPEN TOTAL MASTER1 (2).xlsx"

# Open workbook
wb_obj = openpyxl.load_workbook(path, read_only=True)

# Gets active sheet
sheet_obj = wb_obj['TOTAL LIVE']

# Gets first cell
cell_obj = sheet_obj.cell(row = 27, column = 8)

print(cell_obj.value)
