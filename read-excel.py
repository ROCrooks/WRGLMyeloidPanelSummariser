#Import required modules
import pandas as pd

#Get the Excel file and the list of sheets which it contains
excel = pd.ExcelFile('example-excel.xlsx')
sheets = excel.sheet_names

sheet = excel.parse("Sheet1")
print(sheet)

#Test dataframe
df1 = pd.DataFrame([['a', 'b'], ['c', 'd']],
#index=['row 1', 'row 2'],
#columns=['col 1', 'col 2']
)

df1.to_excel("output.xlsx",
header = False,
index = False
)

#df = pd.read_excel(r'example-excel.xlsx')
#print(df)
