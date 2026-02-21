from os import listdir
import os
import openpyxl
from operator import itemgetter

from traitement_excel import *


currentfolderwin = str(folderChoice())
currentfolder = str(folderconvert(currentfolderwin))

fileInFolderList = (listdir(currentfolder))

filelist = filetest(fileInFolderList)
print("les fichiers retenu sont " + str(filelist))

os.chdir(currentfolderwin)
print(currentfolder)

checkfile(filelist)

allCustomersInFolderList = takeclientlist(filelist)
print(allCustomersInFolderList)

customersdatalist = getcustomersdata(allCustomersInFolderList, filelist)
print(customersdatalist)

for check in customersdatalist:
    for x in check:
        i = 0
        for checkin in x:
            if checkin == None:
                x[i] = 0
            i += 1

print('initialisation terminer')

for toto in customersdatalist:
    get_n = itemgetter(1)
    myList = toto
    myList.sort(key=get_n)

if not os.path.exists(str(currentfolder) + "/factureclients"):
    os.makedirs(str(currentfolder) + "/factureclients")

os.chdir(currentfolderwin + "/factureclients")

for data in customersdatalist:
    wbOutput = openpyxl.Workbook()
    sheet = wbOutput.active
    sheet["A1"] = "Facture"
    sheet["B6"] = "DATE"
    sheet["C6"] = "Maxi"
    sheet["D6"] = "Geant"
    sheet["E6"] = "Miche"
    sheet["F6"] = "Galette"
    sheet["G6"] = "Somun"
    sheet["H6"] = "Marguerite"
    sheet["I6"] = "Pide"

    linetable = 7
    for inDATA in data:
        filenameOutput = str(inDATA[0])
        date = inDATA[1].strftime("%d/%m/%Y")
        maxi = inDATA[2]
        geant = inDATA[3]
        miche = inDATA[4]
        galette = inDATA[5]
        somun = inDATA[6]
        marguerite = inDATA[7]
        pide = inDATA[8]
        sheet["A2"] = str(filenameOutput)
        sheet["B" + str(linetable)] = str(date)
        sheet["C" + str(linetable)] = int(maxi)
        sheet["D" + str(linetable)] = int(geant)
        sheet["E" + str(linetable)] = int(miche)
        sheet["F" + str(linetable)] = int(galette)
        sheet["G" + str(linetable)] = int(somun)
        sheet["H" + str(linetable)] = int(marguerite)
        sheet["I" + str(linetable)] = int(pide)
        linetable += 1

    wbOutput.save(str(filenameOutput + ".xlsx"))