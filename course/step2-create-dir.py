import os
import shutil

target_dir = "/Users/tguehring/Downloads/testtesttest"
excel_sheet = "/Users/tguehring/Projects/private/python-excel-utilities/excelsheets/Sheet1.xlsx"


# Cleanup
try:
    os.rmdir(target_dir)
except:
    pass







def createdirectory(path):
    os.mkdir(path)

#to copy file into target directory
def copyfileintodirectory(filepath, targetdirectory):
    shutil.copy(filepath, targetdirectory)

createdirectory(target_dir)
copyfileintodirectory(excel_sheet, target_dir)