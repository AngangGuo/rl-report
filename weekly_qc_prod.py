import openpyxl
from _datetime import datetime


# return the column number and the title of the column
# parameter: title row
def get_col_title(row):
    n = 0
    col_name = {}
    for name in row[1:]:
        n += 1
        if name:
            col_name[n] = name
    return col_name


# get number from column
def get_number(content):
    # ignore OFF, QC, SHP, NoneType, etc. which can't convert to int
    count = 0
    try:
        count = int(content)
    except:
        count = 0
        pass
    return count


# calculate data in the row
def calc_row(col_title, data, row):
    for col, name in col_title.items():
        data[name]["assets"] += get_number(row[col])
        data[name]["errors"] += get_number(row[col+1])
    return


# show the result
def show(data):
    for name, content in data.items():
        result = name.ljust(10, ' ')
        for k,v in content.items():
            result += " | " + str(v).rjust(3, ' ')
            # print(f"\t{k}: {v}")
        print(result)


# process
# begin, end: the starting, end date to process
# Date format: "2020-05-11"
def process(begin, end):
    # Title is in 2nd line
    TITLE_LINE = 2
    filename = "qc_prod/QC and Productivity-2020.xlsx"
    wb = openpyxl.load_workbook(filename, data_only=True)
    summary_sheet = wb["Summary-Daily"]
    # sheet = wb["QC-Weekly"]
    row_num = 0
    data = {}
    for row in summary_sheet.values:
        row_num += 1
        # process column title
        if row_num == TITLE_LINE:
            col_title = get_col_title(row)
            for name in col_title.values():
                data[name] = {"assets": 0, "errors":0}

        # skid the first 2 title lines
        if row_num <= TITLE_LINE:
            continue

        # process date range
        col_date = row[0].strftime("%Y-%m-%d")
        if begin <= col_date <= end:
            calc_row(col_title, data, row)
    return data


def main():
    result = process("2020-05-18", "2020-05-22")
    show(result)


if __name__ == "__main__":
    main()