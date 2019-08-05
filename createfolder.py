import os
import datetime
from os import listdir
from os.path import isfile, join

def createfolder():
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

    #Make new directory and put the labelled file in it
    path = "Myeloid" + currentdate

    os.mkdir(path)

    aidlabel = '"PLACE FILES HERE"'
    touchcommand = "touch " + path + "/" + aidlabel
    os.system(touchcommand)

    #Specify the name of your file manager here
    filemanager = "nautilus"

    #Open folder onto screen
    folderopencommand = filemanager + " " + path

    os.system(folderopencommand)
