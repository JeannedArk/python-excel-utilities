
import os
import shutil
import openpyxl


RESULT_LOCATION = "/Users/tguehring/Downloads/testtesttest"

DATA = "/Users/tguehring/Projects/private/python-excel-utilities/course/Region.xlsx"
REGION_COL_IDX = 1
CRE_ID_COL_IDX = 2


def copy_file_into_dir(source_file_path: str, target_location: str) -> str:
    """
    To copy file into target directory
    """
    return shutil.copy(source_file_path, target_location)


def get_regions_from_excel_sheet(excel_sheet_path: str):
    workbook = openpyxl.load_workbook(excel_sheet_path)
    ws = workbook['Sheet1']
    max_row = ws.max_row
    print(f"Max row: {max_row}")
    region_columns = []
    # Iterate through all the rows to get the regions
    # Start at row 2 to exclude the header 'Region'
    for i in range(2, max_row + 1):
        region = ws.cell(row=i, column=REGION_COL_IDX).value
        print(f"Discovered region: {region}")
        region_columns.append(region)
    # Create a set without duplicates
    return set(region_columns)


def sanitize_excel_sheet(file_path: str, region: str):
    """
    To remove all rows that do not belong to the region
    """
    print("Sanitizing file: " + file_path)

    workbook = openpyxl.load_workbook(file_path)
    ws = workbook['Sheet1']
    max_row = ws.max_row
    # Iterate through all the rows to get the regions
    # Start at row 2 to exclude the header 'Region'
    i = 2
    while i <= max_row:
        row_region = ws.cell(row=i, column=REGION_COL_IDX).value
        print(f"Row {i} has region {row_region}")
        if row_region is None or row_region == "":
            break
        elif row_region != region:
            ws.delete_rows(i)
        else:
            i += 1

    workbook.save(file_path)


def setup_region_folders(regions, target_excel_sheet: str):
    """
    To create folders for each region and copy the sanitized template file into it
    """
    for region in regions:
        dir_path = os.path.join(RESULT_LOCATION, region)
        os.makedirs(dir_path, exist_ok=True)
        print("Created folder: ", dir_path)
        new_region_file_path = copy_file_into_dir(target_excel_sheet, dir_path)
        sanitize_excel_sheet(new_region_file_path, region)


region_columns = get_regions_from_excel_sheet(DATA)
print("Regions: " + str(region_columns))

setup_region_folders(region_columns, DATA)