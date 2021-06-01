# Writing to an excel 
# sheet using Python
import time
import xlwt
from xlwt import Workbook
import xlrd
from datetime import datetime, timedelta
from datetime import date
import random



def createReport(title, IMPORTNAMES, IMPORTVIEWS, IMPORTHOVERS, IMPORTCLICKS, addEmail, addLink, addTweets):
    
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

    sheet1.write(5, 0, "Banner Advertisement", style)
    sheet1.write(5, 1, "Views", style2)
    sheet1.write(5, 2, "Hovers", style2)
    sheet1.write(5, 3, "Clicks", style2)

    totalViews = 0
    totalHovers = 0
    totalClicks = 0

    j = 6
    for i in IMPORTNAMES:
        sheet1.write(j, 0, i, style)
        j = j+1

    j = 6
    for i in IMPORTVIEWS:
        totalViews += int(i)
        sheet1.write(j, 1, i, style2)
        j = j+1

    j = 6
    for i in IMPORTHOVERS:
        totalHovers += int(i)
        sheet1.write(j, 2, i, style2)
        j = j+1

    j = 6
    for i in IMPORTCLICKS:
        totalClicks += int(i)
        sheet1.write(j, 3, i, style2)
        j = j+1
    
    x = len(IMPORTNAMES)+8

    sheet1.write(x, 0, "TOTAL:", style)
    sheet1.write(x, 1, xlwt.Formula("SUM(B8:B{})".format(x-3)), style2)
    sheet1.write(x, 2, xlwt.Formula("SUM(C8:C{})".format(x-3)), style2)
    sheet1.write(x, 3, xlwt.Formula("SUM(D8:D{})".format(x-3)), style2)

    if(addEmail == True):
        addEmailToSheet(sheet1, style, style2, x)
        x = len(IMPORTNAMES)+8
    if(addLink == True):
        addLinkToSheet(sheet1, style, style2, x)
        x = len(IMPORTNAMES)+8
    if(addTweets == True):
        addTweetsToSheet(sheet1, style, style2, x)
        x = len(IMPORTNAMES)+8

    fileName = "{} {}".format(title,date)
    wb.save("C:\Jared\EmpireReport\Reports\\Automated\\" + fileName +".xls")

    totals = [totalViews, totalHovers, totalClicks]

    return totals

def addEmailToSheet(sheet1, style, style2, x):
    
    sheet1.write(x+2, 0, "Email Blast w/ sponsored message", style)
    sheet1.write(x+2, 1, "Impressions", style2)
    sheet1.write(x+2, 3, "Clicks", style2)
    date = datetime.date(datetime.now())

    sheet1.write(x+7, 0, "{}".format(date), style)
    sheet1.write(x+6, 0, "{}".format(date-timedelta(days = 1)), style)
    sheet1.write(x+5, 0, "{}".format(date-timedelta(days = 2)), style)
    sheet1.write(x+4, 0, "{}".format(date-timedelta(days = 3)), style)
    sheet1.write(x+3, 0, "{}".format(date-timedelta(days = 4)), style)

    sheet1.write(x+7, 1, 0, style2)
    sheet1.write(x+6, 1, 0, style2)
    sheet1.write(x+5, 1, 0, style2)
    sheet1.write(x+4, 1, 0, style2)
    sheet1.write(x+3, 1, 0, style2)

    sheet1.write(x+7, 3, 0, style2)
    sheet1.write(x+6, 3, 0, style2)
    sheet1.write(x+5, 3, 0, style2)
    sheet1.write(x+4, 3, 0, style2)
    sheet1.write(x+3, 3, 0, style2)

    sheet1.write(x+9, 0, "TOTAL:", style)
    sheet1.write(x+9, 1, xlwt.Formula("SUM(B{}:B{})".format(x+3,x+7)), style2)
    #sheet1.write(x+9, 2, xlwt.Formula("SUM(C{}:C{})".format(x+3,x+7)), style2)
    sheet1.write(x+9, 3, xlwt.Formula("SUM(D{}:D{})".format(x+3,x+7)), style2)

def addLinkToSheet(sheet1, style, style2, x):
    sheet1.write(x+11, 0, "Sponsored Link", style)
    sheet1.write(x+11, 1, "Impressions", style)
    sheet1.write(x+11, 3, "Clicks", style)

    sheet1.write(x+12, 0, "Title of link", style)

    sheet1.write(x+12, 1, 0, style2)

    sheet1.write(x+12, 3, 0, style2)

    sheet1.write(x+14, 0, "TOTAL:", style)
    sheet1.write(x+14, 1, xlwt.Formula("SUM(B{}:B{})".format(x+12,x+12)), style2)
    #sheet1.write(x+9, 2, xlwt.Formula("SUM(C{}:C{})".format(x+3,x+7)), style2)
    sheet1.write(x+14, 3, xlwt.Formula("SUM(D{}:D{})".format(x+12,x+12)), style2)

def addTweetsToSheet(sheet1, style, style2, x):
    sheet1.write(x+16, 0, "Tweets", style)
    sheet1.write(x+16, 1, "Impressions", style)
    sheet1.write(x+16, 2, "Engagements", style)
    sheet1.write(x+16, 3, "URL Clicks", style)

    sheet1.write(x+17, 0, "Tweet Date", style)
    sheet1.write(x+18, 0, "Tweet Date", style)
    sheet1.write(x+19, 0, "Tweet Date", style)
    sheet1.write(x+20, 0, "Tweet Date", style)
    sheet1.write(x+21, 0, "Tweet Date", style)

    sheet1.write(x+17, 1, 0, style2)
    sheet1.write(x+18, 1, 0, style2)
    sheet1.write(x+19, 1, 0, style2)
    sheet1.write(x+20, 1, 0, style2)
    sheet1.write(x+21, 1, 0, style2)

    sheet1.write(x+17, 2, 0, style2)
    sheet1.write(x+18, 2, 0, style2)
    sheet1.write(x+19, 2, 0, style2)
    sheet1.write(x+20, 2, 0, style2)
    sheet1.write(x+21, 2, 0, style2)

    sheet1.write(x+17, 3, 0, style2)
    sheet1.write(x+18, 3, 0, style2)
    sheet1.write(x+19, 3, 0, style2)
    sheet1.write(x+20, 3, 0, style2)
    sheet1.write(x+21, 3, 0, style2)

    sheet1.write(x+17, 4, "Permalink", style)
    sheet1.write(x+18, 4, "Permalink", style)
    sheet1.write(x+19, 4, "Permalink", style)
    sheet1.write(x+20, 4, "Permalink", style)
    sheet1.write(x+21, 4, "Permalink", style)

    sheet1.write(x+23, 0, "TOTAL:", style)
    sheet1.write(x+23, 1, xlwt.Formula("SUM(B{}:B{})".format(x+17,x+21)), style2)
    sheet1.write(x+23, 2, xlwt.Formula("SUM(C{}:C{})".format(x+17,x+21)), style2)
    sheet1.write(x+23, 3, xlwt.Formula("SUM(D{}:D{})".format(x+17,x+21)), style2)

