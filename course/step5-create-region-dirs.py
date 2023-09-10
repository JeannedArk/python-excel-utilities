import os
import shutil

target_dir = "/Users/tguehring/Downloads/testtesttest"

excel_sheet = "/Users/tguehring/Projects/private/python-excel-utilities/excelsheets/Sheet1.xlsx"
regions = ['EMEA', 'APAC', 'CH', 'Americas']

# Cleanup
try:
    os.rmdir(target_dir)
except:
    pass









# Create the target dir, create within that dir region dirs, copy the excel_sheet into the region dirs
# here you go...oh lord save me 

def createdirectory(path):
    os.mkdir(path)

#to copy file into target directory
def copyfileintodirectory(filepath, target_directory):
    shutil.copy(filepath, target_directory)

# regional folders
for region in regions:
    # use here os.path.join instead of manually concat paths
    regionaldir = os.path.join(target_dir, region)
    os.makedirs(regionaldir)
    copyfileintodirectory(excel_sheet, regionaldir)