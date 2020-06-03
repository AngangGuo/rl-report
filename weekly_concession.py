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


def show(values):
    n = -1
    result = {}
    for v in values:
        n += 1
        # skip the title line
        if n == 0:
            continue
        s = (v[9].strip().title())[:6]
        if result.get(s) is None:
            result[s] = 0

        result[s] += 1

    # print(f"{result}")
    for k,v in result.items():
        print(f"{k}: {v}")
    print(f"Total: {n}")


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

    show(new_sheet.values)

    try:
        new_wb.save(dst)
    except PermissionError:
        print(f"Sorry, I can't save file {dst}. Please close the file and try again.")


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

    shutil.copyfile(SRC_CONCESSION_FILE, DST_CONCESSION_FILE)

    print(f"Concession Report From {begin} to {end}:")
    main(begin, end, DST_CONCESSION_FILE, WEEKLY_FILE)
