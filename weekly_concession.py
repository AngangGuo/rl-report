import openpyxl
import warnings

def main():
    TITLE_LINE = 2
    filename = "concession/Concession v2.xlsx"

    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(filename, data_only=True)
    # Restore to default
    warnings.simplefilter("default")

    new_wb=openpyxl.Workbook()
    new_sheet=new_wb.active

    # row_num = 0
    # data = {}
    for sheet in wb:
        title = sheet.title
        if title == "Sum" or title == "Weekly":
            continue
        # print(title)

        n = 0
        for row in sheet.values:
            # skip title
            n += 1
            if n < 2:
                continue

            # ignore blank row or non data row
            try:
                col_date = row[1].strftime("%Y-%m-%d")
            except:
                col_date = "2020-01-01"
                pass

            if "2020-05-18" <= col_date <= "2020-05-22":
                new_row=[v for v in row]
                new_row[1]=col_date
                new_row.insert(0,title)
                new_sheet.append(new_row)

    new_wb.save("concession/weekly.xlsx")
# for row in summary_sheet.values:


# new_wb=openpyxl.Workbook()
# new_sheet=new_wb.active
#
#
# new_wb.save("temp.xlsx")

if __name__=="__main__":
    main()
