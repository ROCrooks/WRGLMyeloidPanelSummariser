#Import required modules
import pandas as pd

excel = pd.ExcelFile('example-excel.xlsx')

sheets = excel.sheet_names

print(sheets)

#df = pd.read_excel(r'example-excel.xlsx')
#print(df)
