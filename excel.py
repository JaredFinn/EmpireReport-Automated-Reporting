# Writing to an excel 
# sheet using Python
import time
import xlwt
from xlwt import Workbook
import xlrd
from datetime import datetime
import random


def createReport(title, IMPORTNAMES, IMPORTVIEWS, IMPORTHOVERS, IMPORTCLICKS):
    # Workbook is created
    wb = Workbook()

    # add_sheet is used to create sheet.
    sheet1 = wb.add_sheet('Sheet 1')

    font = xlwt.Font() # Create the Font
    font.name = 'Calibri'
    font.height = 220
    style = xlwt.XFStyle() # Create the Style
    style.font = font # Apply the Font to the Style
    style2 = xlwt.XFStyle()
    style2.font = font
    style2.num_format_str = "#,##0"
    al = xlwt.Alignment()
    al.horz = al.HORZ_CENTER
    style2.alignment = al

    date = datetime.date(datetime.now())
    ER = "Empire Report Stats {}".format(date)

    sheet1.write(0, 0, title, style)
    sheet1.write(1, 0, ER, style)

    sheet1.write(6, 0, "Banner Advertisement", style)
    sheet1.write(6, 1, "Views", style2)
    sheet1.write(6, 2, "Hovers", style2)
    sheet1.write(6, 3, "Clicks", style2)

    totalViews = 0
    totalHovers = 0
    totalClicks = 0

    j = 7
    for i in IMPORTNAMES:
        sheet1.write(j, 0, i, style)
        j = j+1

    j = 7
    for i in IMPORTVIEWS:
        totalViews += int(i)
        sheet1.write(j, 1, i, style2)
        j = j+1

    j = 7
    for i in IMPORTHOVERS:
        totalHovers += int(i)
        sheet1.write(j, 2, i, style2)
        j = j+1

    j = 7
    for i in IMPORTCLICKS:
        totalClicks += int(i)
        sheet1.write(j, 3, i, style2)
        j = j+1
    
    x = len(IMPORTNAMES)+10

    sheet1.write(x, 0, "TOTAL:", style)
    sheet1.write(x, 1, xlwt.Formula("SUM(B8:B{})".format(x-3)), style2)
    sheet1.write(x, 2, xlwt.Formula("SUM(C8:C{})".format(x-3)), style2)
    sheet1.write(x, 3, xlwt.Formula("SUM(D8:D{})".format(x-3)), style2)
#
    fileName = "{} {}".format(title,date)
    wb.save("C:\Jared\EmpireReport\Reports\\Automated\\" + fileName +".xls")
    
    totals = [totalViews, totalHovers, totalClicks]

    return totals