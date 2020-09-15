import openpyxl
import datetime
import sys


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
        data[name]["errors"] += get_number(row[col + 1])
    return


# show the result
def show(begin, end, data):
    print(f"[QC and Production Report] from {begin} to {end}:")
    for name, content in data.items():
        result = name.ljust(10, ' ')
        for k, v in content.items():
            result += " | " + str(v).rjust(3, ' ')
        print(result)


# process
# begin, end: the starting, end date to process
# Date format: "2020-05-11"
def process(begin, end):
    # Title is in 2nd line
    TITLE_LINE = 23
    filename = "C:/Users/caguoa00/OneDrive - Ingram Micro/Work/QC and Productivity-2020.xlsx"
    data = {}
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
    except PermissionError:
        print(f"Please close this file and try again: \n{filename}")
        return data

    summary_sheet = wb["Summary-Daily"]
    # sheet = wb["QC-Weekly"]
    row_num = 0
    for row in summary_sheet.values:
        row_num += 1
        # process column title
        if row_num == TITLE_LINE:
            col_title = get_col_title(row)
            for name in col_title.values():
                data[name] = {"assets": 0, "errors": 0}

        # skid the first 2 title lines
        if row_num <= TITLE_LINE:
            continue

        # process date range
        col_date = row[0].strftime("%Y-%m-%d")
        if begin <= col_date <= end:
            calc_row(col_title, data, row)
    return data


# def main():
#     result = process("2020-05-25", "2020-05-29")
#     show(result)
def last_week():
    today = datetime.datetime.now()

    weekday = today.weekday()
    last_monday = today + datetime.timedelta(days=(-7 - weekday))
    last_saturday = today + datetime.timedelta(days=(-2 - weekday))

    return last_monday.strftime("%Y-%m-%d"), last_saturday.strftime("%Y-%m-%d")


if __name__ == "__main__":
    begin = ""
    end = ""
    if len(sys.argv) == 3:
        begin = sys.argv[1]
        end = sys.argv[2]
    else:
        begin, end = last_week()

    result = process(begin, end)
    if result != {}:
        show(begin, end, result)
