###########################
##   Author: Tom Prins   ##
###########################

import translators as ts
import pandas as pd
from tqdm import tqdm

excelExtension = ".xlsx"
fileName = input("Enter the file name: ")
if excelExtension not in fileName:
    fileName = fileName+excelExtension
excelFile = pd.read_excel(fileName)

exportFileName = input("What should the output file be called?\n")
columnName = input("Enter the column name, that needs to be translated?\n")

print("1. Translate from English")
print("2. Translate from German")
translateOption = input("Type 1 or 2: ")
if translateOption == 1:
    fromLanguage = "en"
else:
    fromLanguage = "de"

notTranslatedTextList = []
translatedTextList = []


# Make an array with all the values of fabrikantnummer
def fileToArray(excelFile):
    # search excel file for 'fabrikantnummer', if not found, take the first column
    try:
        excelFile = excelFile[columnName].tolist()
    except(KeyError):
        excelFile = excelFile.iloc[:, 0].tolist()

    excelFile = [x for x in excelFile if str(x) != 'nan']

    # Retrun a List of all the fabrikantnummers of the file
    return excelFile

def translateArray(notTranslatedTextList):

    for index in tqdm(range(len(notTranslatedTextList))):
        translatedTextList.append(ts.google(str(notTranslatedTextList[index]), from_language=fromLanguage, to_language="nl"))
    
    # Export the arrays to an excelfile
    translation = pd.DataFrame({'Translated': translatedTextList})
    writer = pd.ExcelWriter(str(exportFileName+".xlsx"), engine='xlsxwriter')
    translation.to_excel(writer, sheet_name="Sheet1", header=False, index=False)
    writer.save()

notTranslatedTextList = fileToArray(excelFile)
translateArray(notTranslatedTextList)


