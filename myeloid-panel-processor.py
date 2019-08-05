#Import required modules
import pandas as pd

#Container to hold errors
errors = []

#Import the excel file
try:
    excelfile = pd.ExcelFile('example-excel.xlsx')
except:
    errors.append("The file does not appear to be a valid Excel file")

#Get all the sheet names from the Excel file
try:
    from readexcelfile import getexcelwnumbersheetnames
    sheetnames = getexcelwnumbersheetnames(excelfile)
except:
    errors.append("The Excel file does not appear to have any readable sheets")

#Print errors that were collected
if len(errors) > 0:
    for error in errors:
        print("Error: " + error)
