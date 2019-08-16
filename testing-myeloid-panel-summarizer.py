#Use pandas
import pandas as pd
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
    workingpath = createfolder(currentdate,"Testing")
except:
    errors.append("Cannot create folder for testing")

#Prompt for user to place Excel file(s) into folder
input("Copy your files for testing pre and post update into the folder and press Enter to continue...")

#Select the excel file from the folder
try:
    from selectexcel import selectexcel
    excelfile1 = selectexcel(workingpath,1)
    excelfile2 = selectexcel(workingpath,2)
except:
    errors.append("Cannot access any Excel files in the directory")

#Read the Excel files
try:
    excel1 = pd.ExcelFile(excelfile1)
    excel2 = pd.ExcelFile(excelfile2)
except:
    errors.append("Cannot access any Excel files in the directory")

#Get all the sheet names from the Excel file
try:
    from readexcelfile import getexcelwnumbersheetnames
    sheetnames1 = getexcelwnumbersheetnames(excel1)
    sheetnames2 = getexcelwnumbersheetnames(excel2)
except:
    errors.append("The Excel file does not appear to have any readable sheets")

print(sheetnames1)
print(sheetnames2)

#difference = df1[df1!=df2]
#differencesum = difference.sum()
#differencesum = differencesum.sum()
#print(differencesum)

#if difference == none:
#    print("The Excel Files are identical!")
