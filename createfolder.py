import os
from os import listdir
from os.path import isfile, join, isdir
from shutil import rmtree
import subprocess

def createfolder(currentdate,name):
    #Make new directory and put the labelled file in it
    path = "output/" + name + currentdate

    #Check if folder exists
    if os.path.isdir(path) == True:
        #Ask if want to overwrite, do not accept anything that isn't y or n
        overwrite = ""
        while overwrite != 'y' and overwrite != 'n':
            overwrite = input("A myeloid panel folder has already been created today. Do you want to overwrite the current one? (y/n) ")

        #If yes delete
        if overwrite == "y":
            rmtree(path)
        elif overwrite == "n":
            #Append a flag at the end of the folder name
            appendno = 1
            path = path + "-" + str(appendno)
            #Incremenent flag number until the folder name doesn't exist
            while os.path.isdir(path) == True:
                appendno = appendno+1
                path = path + "-" + str(appendno)

        #If no create new folder name

    os.mkdir(path)

    #Create a file saying to place the file in the open directory
    aidlabel = '"PLACE FILES HERE"'
    #touchcommand = "touch " + path + "/" + aidlabel
    #os.system(touchcommand)
    touchcommand = path + "/" + aidlabel
    open(touchcommand,'a').close()

    #Specify the name of your file manager here
    filemanager = "nautilus"

    #Open folder onto screen
    folderopencommand = filemanager + ' "' + path + '"'

    os.system(folderopencommand)
    #FNULL = open(os.devnull, 'w')
    #subprocess.Popen([filemanager, path],close_fds=True,stout=FNULL,stderr=subprocess.STDOUT)
    #os.startfile(folderopencommand)

    return path
