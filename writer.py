import xlsxwriter


def write_headers(worksheet, format):
    headers = ["Name", "Date", "Hours", "Description"]
    for i, header in enumerate(headers):
        worksheet.write(0, i, header, format)


def write_data(worksheet, data, format):
    row, col = 1, 0
    for name, date, hours, description in data:
        worksheet.write(row, col, name, format)
        worksheet.write(row, col+1, date, format)
        worksheet.write(row, col+2, hours, format)
        worksheet.write(row, col+3, description, format)
        row, col = row + 1, 0


def write_raport(data, filename):    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    
    column_widths = (20, 10, 5, 60)
    for i, width in enumerate(column_widths):
        worksheet.set_column(i, i, width)

    default_format = workbook.add_format({"align": "center"})
    header_format = workbook.add_format({"bold": True, "align": "center"})

    write_headers(worksheet, header_format)
    write_data(worksheet, data, default_format)

    workbook.close()
