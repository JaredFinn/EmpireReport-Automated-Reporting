# Writing to an excel 
# sheet using Python
import time
import xlwt
from xlwt import Workbook
import xlrd
from datetime import datetime
import random

def createReport(title):
    # Workbook is created
    wb = Workbook()

    # add_sheet is used to create sheet.
    sheet1 = wb.add_sheet('Sheet 1')

    font = xlwt.Font() # Create the Font
    font.name = 'Calibri'
    font.height = 220
    style = xlwt.XFStyle() # Create the Style
    style.font = font # Apply the Font to the Style
    style.num_format_str = "#,##0"
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
    sheet1.write(7, 0, "Ad1", style)
    sheet1.write(8, 0, "Ad2", style)
    sheet1.write(9, 0, "Ad3", style)
    sheet1.write(10, 0, "Ad4", style)

    sheet1.write(6, 1, "Views", style2)
    val1 = random.randrange(100, 100000, 17)
    sheet1.write(7, 1, val1, style2)
    val2 = random.randrange(100, 100000, 17)
    sheet1.write(8, 1, val2, style2)
    val3 = random.randrange(100, 100000, 17)
    sheet1.write(9, 1, val3, style2)
    val4 = random.randrange(100, 100000, 17)
    sheet1.write(10, 1, val4, style2)

    sheet1.write(6, 2, "Hovers", style2)
    val5 = random.randrange(100, 100000, 17)
    sheet1.write(7, 2, val5, style2)
    val6 = random.randrange(100, 100000, 17)
    sheet1.write(8, 2, val6, style2)
    val7 = random.randrange(100, 100000, 17)
    sheet1.write(9, 2, val7, style2)
    val8 = random.randrange(100, 100000, 17)
    sheet1.write(10, 2, val8, style2)

    sheet1.write(6, 3, "Clicks", style2)
    val9 = random.randrange(100, 100000, 17)
    sheet1.write(7, 3, val9, style2)
    val10 = random.randrange(100, 100000, 17)
    sheet1.write(8, 3, val10, style2)
    val11 = random.randrange(100, 100000, 17)
    sheet1.write(9, 3, val11, style2)
    val12 = random.randrange(100, 100000, 17)
    sheet1.write(10, 3, val12, style2)

    sheet1.write(13, 0, "TOTAL:", style)
    sheet1.write(13, 1, xlwt.Formula("SUM(B8:B11)"), style2)
    sheet1.write(13, 2, xlwt.Formula("SUM(C8:C11)"), style2)
    sheet1.write(13, 3, xlwt.Formula("SUM(D8:D11)"), style2)

    fileName = "{} {}".format(title,date)
    wb.save(fileName+".xls")

    totalViews = val1+val2+val3+val4
    totalHovers =val5+val6+val7+val8
    totalClicks =val9+val10+val11+val12

    totals = [totalViews, totalHovers, totalClicks]

    return totals

