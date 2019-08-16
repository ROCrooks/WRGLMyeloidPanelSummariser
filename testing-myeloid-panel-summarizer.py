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

mismatches = []

#Look for mismatches Excel1 into Excel2
for sheetname in sheetnames1:
    #Extract Excel sheet as a dataframe
    df1 = excel1.parse(sheetname)

    #See if a sheet in excel2 exists
    try:
        df2 = excel2.parse(sheetname)
        comparesheets = True
    except:
        mismatches.append(sheetname + " is not in both spreadsheet files.")
        comparesheets = False

    try:
        if comparesheets is True:
            #Calculate the difference between the Excel sheets
            difference = df1[df1!=df2]
            differencesum = difference.sum()
            differencesum = differencesum.sum()

            #Append as mismatch if there is a difference
            if differencesum > 0:
                mismatches.append(sheetname + " is not the same in both files.")
    except:
        errors.append("There is an error in comparing the Excel sheets named " + sheetname)
#Look for mismatches Excel1 into Excel2
for sheetname in sheetnames2:
    #Extract Excel sheet as a dataframe
    df1 = excel2.parse(sheetname)

    #See if a sheet in excel2 exists
    try:
        df2 = excel1.parse(sheetname)
        comparesheets = True
    except:
        mismatches.append(sheetname + " is not in both spreadsheet files.")
        comparesheets = False

    try:
        if comparesheets is True:
            #Calculate the difference between the Excel sheets
            difference = df1[df1!=df2]
            differencesum = difference.sum()
            differencesum = differencesum.sum()

            #Append as mismatch if there is a difference
            if differencesum > 0:
                mismatches.append(sheetname + " is not the same in both files.")
    except:
        errors.append("There is an error in comparing the Excel sheets named " + sheetname)


#Print errors that were collected
if len(mismatches) > 0:
    print("Mismatches between the files have been identified")
    for mismatch in mismatches:
        print("Mismatch: " + mismatch)
else:
    print("The files are identical!")

#Print errors that were collected
if len(errors) > 0:
    for error in errors:
        print("Error: " + error)

#difference = df1[df1!=df2]
#differencesum = difference.sum()
#differencesum = differencesum.sum()
#print(differencesum)

#if difference == none:
#    print("The Excel Files are identical!")
