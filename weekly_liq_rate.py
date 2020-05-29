import openpyxl
import os
import warnings
from datetime import datetime

# Sheet name of InventoryAllFields.xlsx
SHEET = "Sheet"
# site: FG(CLIENT)
SITE = "Site"
# Sun, Lisa
TESTED_BY = "Tested By"


# return the SITE and TESTED_BY column numbers
def get_col_numbers(list):
    site = -1
    tested_by = -1
    n = -1
    for name in list:
        n += 1
        if name.strip() == SITE:
            site = n
            continue
        if name.strip() == TESTED_BY:
            tested_by = n
            continue
    return site, tested_by


# show the result
def show(data):
    print(f"\nProcessed on {datetime.now().date()}:")
    sum = 0
    for name, data in data.items():
        last_name, first_name = name.split(',', 1)
        print(f"{first_name} {last_name}:")
        for k, v in data.items():
            sum += v
            print(f"\t{k}:{v}")
    print(f"Total: {sum}")


# process the file and save the result into liq_rate
# filename: InventoryAll.xlsx
def calculate(data, filename):
    # Fix: Use 'warnings.simplefilter("ignore")' to suppress the following warning message
    # ------------------
    # C:\Users\caguoa00\PycharmProjects\learn\venv\lib\site - packages\openpyxl\styles\stylesheet.py: 214:
    # UserWarning: Workbook contains no default style, apply openpyxl's default;
    # warn("Workbook contains no default style, apply openpyxl's default")
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(filename)
    # Restore to default
    warnings.simplefilter("default")
    sheet = wb[SHEET]

    # current row number to process
    row_no = 0
    # Site column number
    site_col = -1
    # Tested By column number
    tested_by_col = -1
    for row in sheet.values:
        row_no += 1
        # first row contains column titles
        if row_no == 1:
            site_col, tested_by_col = get_col_numbers(row)
            # print(f"site col: {site_col}, tested_by number: {tested_by_col}")
            continue
        # print(row[site_col],row[tested_by_col])

        # check tested by key
        dname = data.get(row[tested_by_col])
        if not dname:
            data[row[tested_by_col]] = {}
        # check site key
        dsite = data[row[tested_by_col]].get(row[site_col])
        if not dsite:
            data[row[tested_by_col]][row[site_col]] = 0
        # add one to tested by/site pair
        data[row[tested_by_col]][row[site_col]] += 1


# main function
def process(folder):
    # hold the result for all the files in the folder
    result = {}

    for filename in os.listdir(folder):
        if filename.lower().endswith('xlsx'):
            calculate(result, os.path.join(folder, filename))

    return result


def main():
    # directory of all the Inventory All files
    files_dir = "liq_rate"
    result = process(files_dir)
    show(result)

# run the program
if __name__ == "__main__":
    main()
