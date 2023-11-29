# This python program will help to automate some of the steps required to perform the weekly comm's cell report

# Imports Libraries/Dependencies
import openpyxl
import pandas as pd
from pathlib import Path

# Path of excel sheet and open workbook
path = Path(r"C:\Users\U360269\Desktop\Codes\CommCellAutomate\SPEN TOTAL MASTER1 (2).xlsx")
sheet = ["TOTAL LIVE"]

# Pandas read excel
df = pd.read_excel(path, engine='openpyxl', sheet_name = sheet)

# Print DF to debug
# print(df)
