#Import required modules
import pandas as pd

import os
import datetime
from os import listdir
from os.path import isfile, join


#Container to hold errors
errors = []

#Create the folder for analysing the batch
try:
    from createfolder import createfolder
    workingpath = createfolder()
except:
    errors.append("Cannot create folder")

#Prompt for user to place Excel file(s) into folder
input("Copy your input file into the folder and press Enter to continue...")

#Select the excel file from the folder
from selectexcel import selectexcel
excelfile = selectexcel(workingpath)

print(excelfile)

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
