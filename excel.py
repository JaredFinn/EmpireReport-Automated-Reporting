# Writing to an excel 
# sheet using Python
import xlwt
from xlwt import Workbook
from datetime import datetime

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
    style2 = xlwt.XFStyle()
    style2.font = font
    al = xlwt.Alignment()
    al.horz = al.HORZ_CENTER
    style2.alignment = al

    date = datetime.date(datetime.now())
    ER = "Empire Report Stats {}".format(date)
    val = 0


    sheet1.write(0, 0, title, style)
    sheet1.write(1, 0, ER, style)

    sheet1.write(6, 0, "Banner Advertisement", style)
    sheet1.write(7, 0, "Ad1", style)
    sheet1.write(8, 0, "Ad2", style)
    sheet1.write(9, 0, "Ad3", style)
    sheet1.write(10, 0, "Ad4", style)

    sheet1.write(6, 1, "Views", style2)
    sheet1.write(7, 1, val, style2)
    sheet1.write(8, 1, val, style2)
    sheet1.write(9, 1, val, style2)
    sheet1.write(10, 1, val, style2)

    sheet1.write(6, 2, "Hovers", style2)
    sheet1.write(7, 2, val, style2)
    sheet1.write(8, 2, val, style2)
    sheet1.write(9, 2, val, style2)
    sheet1.write(10, 2, val, style2)

    sheet1.write(6, 3, "Clicks", style2)
    sheet1.write(7, 3, val, style2)
    sheet1.write(8, 3, val, style2)
    sheet1.write(9, 3, val, style2)
    sheet1.write(10, 3, val, style2)


    sheet1.write(13, 0, "TOTAL:", style)
    totalViews = sheet1.write(13, 1, val, style2)
    totalHovers = sheet1.write(13, 2, val, style2)
    totalClicks = sheet1.write(13, 3, val, style2)

    fileName = "{} {}".format(title,date)
    wb.save(fileName+".xls")

    totals = [totalViews, totalHovers, totalClicks]

    return totals
