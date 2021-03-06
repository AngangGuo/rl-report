import openpyxl
import warnings
import datetime
import sys
import shutil


# return the first and last day of last week
def last_week():
    today = datetime.datetime.now()

    weekday = today.weekday()
    last_monday = today + datetime.timedelta(days=(-7 - weekday))
    last_saturday = today + datetime.timedelta(days=(-2 - weekday))

    return last_monday.strftime("%Y-%m-%d"), last_saturday.strftime("%Y-%m-%d")


def show(begin, end, values):
    # column number for "Possible Issue From" in Weekly sheet
    col = 9

    row = -1
    result = {}
    print(f"Concession Report From {begin} to {end}:")
    for v in values:
        row += 1
        # skip the title row
        if row == 0:
            continue
        issue_from = v[col]
        # fix: TypeError: 'NoneType' object is not subscriptable
        try:
            # Keep the first 6 characters only for easy sum up
            # customer == customer / Ingram == Customer/Amazon ...
            issue_from = issue_from[:6].strip().title()
        except TypeError:
            print(issue_from)
            continue

        if result.get(issue_from) is None:
            result[issue_from] = 0

        result[issue_from] += 1

    # print(f"{result}")
    for k, v in result.items():
        print(f"{k}: {v}")
    print(f"Total: {row}")


def main(begin, end, src, dst):
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(src, data_only=True)
    # Restore to default
    warnings.simplefilter("default")

    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active

    # write title to the first line of the new file
    writen = False
    for sheet in wb:
        title = sheet.title
        if title == "Sum" or title == "Weekly":
            continue

        n = 0
        for row in sheet.values:
            n += 1
            # write the first line as title line
            if n == 1 and not writen:
                title_row = list(row[:])
                title_row.insert(0, "Tester")
                new_sheet.append(title_row)
                writen = True

            if n < 2:
                continue

            # ignore blank row or non data row
            try:
                col_date = row[1].strftime("%Y-%m-%d")
            except:
                col_date = "2020-01-01"
                pass

            # if "2020-05-18" <= col_date <= "2020-05-22":
            if begin <= col_date <= end:
                new_row = [v for v in row]
                new_row[1] = col_date
                # tester in first line
                new_row.insert(0, title)
                new_sheet.append(new_row)

    try:
        new_wb.save(dst)
    except PermissionError:
        print(f"Sorry, I can't save file {dst}. Please close the file and try again.")

    return new_sheet.values
    # show(begin, end, new_sheet.values)


# Main Function
if __name__ == "__main__":
    # Original file
    SRC_CONCESSION_FILE = "C:/Users/caguoa00/OneDrive - Ingram Micro/Work/Concession v2.xlsx"
    DST_CONCESSION_FILE = "concession/Concession v2.xlsx"
    # Weekly file - new
    WEEKLY_FILE = "concession/weekly.xlsx"

    begin = ""
    end = ""
    if len(sys.argv) == 3:
        begin = sys.argv[1]
        end = sys.argv[2]
    else:
        begin, end = last_week()

    # Work on the local copy of the file to prevent permission error when others open the file
    shutil.copyfile(SRC_CONCESSION_FILE, DST_CONCESSION_FILE)

    values = main(begin, end, DST_CONCESSION_FILE, WEEKLY_FILE)
    show(begin, end, values)
