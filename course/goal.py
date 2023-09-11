
import os
import shutil
import openpyxl


RESULT_LOCATION = "/Users/tguehring/Downloads/testtesttest"

DATA = "/Users/tguehring/Projects/private/python-excel-utilities/course/Region.xlsx"
REGION_COL_IDX = 1
CRE_ID_COL_IDX = 2

SAMPLE_TEMPLATE = "/Users/tguehring/Projects/private/python-excel-utilities/excelsheets/Sheet1.xlsx"


def copy_file_into_dir(source_file_path, target_location):
    """
    To copy file into target directory
    """
    shutil.copy(source_file_path, target_location)


def get_regions_from_excel_sheet(excel_sheet_path: str):
    workbook = openpyxl.load_workbook(excel_sheet_path)
    ws = workbook['Sheet1']
    max_row = ws.max_row
    region_columns = []
    # Iterate through all the rows to get the regions
    # Start at row 2 to exclude the header 'Region'
    for i in range(2, max_row):
        region = ws.cell(row=i, column=REGION_COL_IDX).value
        print(f"Discovered region: {region}")
        region_columns.append(region)
    # Create a set without duplicates
    return set(region_columns)


def get_credids_from_excel_sheet(excel_sheet_path: str):
    workbook = openpyxl.load_workbook(excel_sheet_path)
    ws = workbook['Sheet1']
    max_row = ws.max_row
    creid_columns = []
    # Iterate through all the rows to get the cre ids
    # Start at row 3 to exclude the header 'CRE ID'
    for i in range(3, max_row):
        cre_id = ws.cell(row=i, column=CRE_ID_COL_IDX).value
        print(f"Discovered CRE ID: {cre_id}")
        creid_columns.append(cre_id)
    # Create a set without duplicates
    return set(creid_columns)


def setup_region_folders(regions):
    """
    To create folders for each region
    """
    for region in regions:
        dir_path = os.path.join(RESULT_LOCATION, region)
        os.makedirs(dir_path)
        print("Created folder: ", dir_path)
        copy_file_into_dir(SAMPLE_TEMPLATE, dir_path)


region_columns = get_regions_from_excel_sheet(DATA)
print("Regions: " + str(region_columns))

setup_region_folders(region_columns)