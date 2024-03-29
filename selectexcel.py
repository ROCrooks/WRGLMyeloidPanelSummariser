#Import required modules
from os import listdir
from os.path import isfile, join
import re

def selectexcel(path,usecase):
    #Get file names
    fileslist = [f for f in listdir(path) if isfile(join(path, f))]

    #Filter file names to only return Excel files
    xlfiles = []
    for file in fileslist:
        if file.endswith(".xlsx") == True:
            xlfiles.append(file)

    #Prompt for selecting Excel file
    if len(xlfiles) > 0:
        filenumber = 1

        #List all excel files and press enter
        for xlfile in xlfiles:
            print(str(filenumber) + ": " + xlfile)
            filenumber = filenumber+1
        if usecase == "only":
            choosenumber = input("Please type the number of the Excel file and press Enter: ")
        elif usecase == 1:
            choosenumber = input("Please type the number of the 1st Excel file and press Enter: ")
        elif usecase == 2:
            choosenumber = input("Please type the number of the 2nd Excel file and press Enter: ")

        #Get the Excel file based on the number specified
        choosenumber = int(choosenumber)-1
        xlfile = xlfiles[choosenumber]
    else:
        xlfile = None

    return xlfile
