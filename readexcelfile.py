#Import required modules
import pandas as pd
import re

def getexcelwnumbersheetnames(excelfile)
    #Get the Excel file and the list of sheets which it contains
    excel = pd.ExcelFile('example-excel.xlsx')
    allsheets = excel.sheet_names

    #Only keep sheet names that match the regular expression
    resultssheets = []
    for sheet in allsheets:
        if (bool(re.match(r"W[0-9][0-9][0-9][0-9][0-9][0-9][0-9]-[0-9][0-9][0-9]",sheet)) == True):
            resultssheets.append(sheet)

    return resultssheets
