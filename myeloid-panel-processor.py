#Import required modules
import pandas as pd
import xlsxwriter

import os
import datetime
from os import listdir
from os.path import isfile, join

import datetime

#Get date and format as ISO 8601
now = datetime.datetime.now()
if now.month < 10:
    month = "0" + str(now.month)
else:
    month = str(now.month)

if now.day < 10:
    day = "0" + str(now.day)
else:
    day = str(now.day)
currentdate = str(now.year) + month + day

#Container to hold errors
errors = []

#Create the folder for analysing the batch
try:
    from createfolder import createfolder
    workingpath = createfolder(currentdate)
except:
    errors.append("Cannot create folder")

#Prompt for user to place Excel file(s) into folder
input("Copy your input file into the folder and press Enter to continue...")

#Select the excel file from the folder
try:
    from selectexcel import selectexcel
    excelfile = selectexcel(workingpath)
except:
    errors.append("Cannot access any Excel files in the directory")

#Import the excel file
try:
    excelfile = pd.ExcelFile(excelfile)
except:
    errors.append("The file does not appear to be a valid Excel file")

#Get all the sheet names from the Excel file
try:
    from readexcelfile import getexcelwnumbersheetnames
    sheetnames = getexcelwnumbersheetnames(excelfile)
except:
    errors.append("The Excel file does not appear to have any readable sheets")

try:
    #Open the workbook
    outputfile = "Myeloid" + currentdate + "/MyeloidCoverageSummary" + currentdate + ".xlsx"
    outputworkbook = xlsxwriter.Workbook(outputfile)

    #Read each sheet in the spreadsheet
    for sheet in sheetnames:
        sheeterror = False
        try:
            #Extract Excel sheet as a dataframe
            df = excelfile.parse(sheet)

            #Make coverage data into a dictionary
            genes = df['ID'].tolist()
            coverage = df['100x'].tolist()
            coveragedata = dict(zip(genes,coverage))
        except:
            errors.append("Cannot read coverage sheet " + sheet)
            sheeterror = True

        #Write the sheet to the output Excel file
        if sheeterror == False:
            try:
                from makeexcelsheet import makeexcelsheet
                makeexcelsheet(outputworkbook,sheet,coveragedata)
            except:
                errors.append("Failure to make sheet " + sheet)

    #Close the output workbook and print success
    outputworkbook.close()
    print("Success! Created the myeloid panel summary spreadsheet!")
except:
    errors.append("Failed to create myeloid panel workbook " + sheet)

#Print errors that were collected
if len(errors) > 0:
    for error in errors:
        print("Error: " + error)
