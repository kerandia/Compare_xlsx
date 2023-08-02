import openpyxl

def compare_excel_files(file1, file2):
    workbook1 = openpyxl.load_workbook(file1)
    workbook2 = openpyxl.load_workbook(file2)

    sheet1 = workbook1.active
    sheet2 = workbook2.active

    rows1 = sheet1.rows
    rows2 = sheet2.rows
    counter = 0
    result = []
    for row1, row2 in zip(rows1, rows2):
        counter += 1
        value1 = row1[0].value
        value2 = row2[0].value
        if value1 == value2:
            result.append("Yes, row{}".format(counter))
        else:
            result.append("No, row {}".format(counter))
    return result

if __name__ == "__main__":
    file1 = "file1.xlsx"
    file2 = "file2.xlsx"

    result = compare_excel_files(file1, file2)
    print(result)
