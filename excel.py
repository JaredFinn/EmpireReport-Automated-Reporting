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
    
    x = len(IMPORTNAMES)+6

    if(addEmail == True | addLink == True | addTweets == True):
        sheet1.write(x, 0, "SUBTOTAL:", style)
        sheet1.write(x, 1, xlwt.Formula("SUM(B7:B{})".format(x)), style2)
        sheet1.write(x, 2, xlwt.Formula("SUM(C7:C{})".format(x)), style2)
        sheet1.write(x, 3, xlwt.Formula("SUM(D7:D{})".format(x)), style2)

    x = x + 1
    if(addEmail == True):
        x = addEmailToSheet(sheet1, style, style2, x)
    if(addLink == True):
        x = addLinkToSheet(sheet1, style, style2, x)
    if(addTweets == True):
        x = addTweetsToSheet(sheet1, style, style2, x)

    if(addEmail == True | addLink == True | addTweets == True):
        sheet1.write(x+2, 0, "GRAND TOTAL:", style)
        sheet1.write(x+2, 1, 0, style2)
        sheet1.write(x+2, 2, 0, style2)
        sheet1.write(x+2, 3, 0, style2)
    else:
        sheet1.write(x+1, 0, "TOTAL:", style)
        sheet1.write(x+1, 1, xlwt.Formula("SUM(B7:B{})".format(x)), style2)
        sheet1.write(x+1, 2, xlwt.Formula("SUM(C7:C{})".format(x)), style2)
        sheet1.write(x+1, 3, xlwt.Formula("SUM(D7:D{})".format(x)), style2)
    


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

    sheet1.write(x+8, 0, "SUBTOTAL:", style)
    sheet1.write(x+8, 1, xlwt.Formula("SUM(B{}:B{})".format(x+4,x+8)), style2)
    #sheet1.write(x+9, 2, xlwt.Formula("SUM(C{}:C{})".format(x+3,x+7)), style2)
    sheet1.write(x+8, 3, xlwt.Formula("SUM(D{}:D{})".format(x+4,x+8)), style2)

    x = x+9
    return x

def addLinkToSheet(sheet1, style, style2, x):
    sheet1.write(x+2, 0, "Sponsored Link", style)
    sheet1.write(x+2, 1, "Impressions", style2)
    sheet1.write(x+2, 3, "Clicks", style2)

    sheet1.write(x+3, 0, "Title of link", style)

    sheet1.write(x+3, 1, 0, style2)

    sheet1.write(x+3, 3, 0, style2)

    #sheet1.write(x+4, 0, "SUBTOTAL:", style)
    #sheet1.write(x+4, 1, xlwt.Formula("SUM(B{}:B{})".format(x+13,x+13)), style2)
    #sheet1.write(x+9, 2, xlwt.Formula("SUM(C{}:C{})".format(x+3,x+7)), style2)
    #sheet1.write(x+4, 3, xlwt.Formula("SUM(D{}:D{})".format(x+13,x+13)), style2)

    x = x+4
    return x

def addTweetsToSheet(sheet1, style, style2, x):
    sheet1.write(x+2, 0, "Tweets", style)
    sheet1.write(x+2, 1, "Impressions", style2)
    sheet1.write(x+2, 2, "Engagements", style2)
    sheet1.write(x+2, 3, "URL Clicks", style2)

    sheet1.write(x+3, 0, "Tweet Date", style)
    sheet1.write(x+4, 0, "Tweet Date", style)
    sheet1.write(x+5, 0, "Tweet Date", style)
    sheet1.write(x+6, 0, "Tweet Date", style)
    sheet1.write(x+7, 0, "Tweet Date", style)

    sheet1.write(x+3, 1, 0, style2)
    sheet1.write(x+4, 1, 0, style2)
    sheet1.write(x+5, 1, 0, style2)
    sheet1.write(x+6, 1, 0, style2)
    sheet1.write(x+7, 1, 0, style2)

    sheet1.write(x+3, 2, 0, style2)
    sheet1.write(x+4, 2, 0, style2)
    sheet1.write(x+5, 2, 0, style2)
    sheet1.write(x+6, 2, 0, style2)
    sheet1.write(x+7, 2, 0, style2)

    sheet1.write(x+3, 3, 0, style2)
    sheet1.write(x+4, 3, 0, style2)
    sheet1.write(x+5, 3, 0, style2)
    sheet1.write(x+6, 3, 0, style2)
    sheet1.write(x+7, 3, 0, style2)

    sheet1.write(x+3, 4, "Permalink", style)
    sheet1.write(x+4, 4, "Permalink", style)
    sheet1.write(x+5, 4, "Permalink", style)
    sheet1.write(x+6, 4, "Permalink", style)
    sheet1.write(x+7, 4, "Permalink", style)

    sheet1.write(x+8, 0, "SUBTOTAL:", style)
    sheet1.write(x+8, 1, xlwt.Formula("SUM(B{}:B{})".format(x+4,x+8)), style2)
    sheet1.write(x+8, 2, xlwt.Formula("SUM(C{}:C{})".format(x+4,x+8)), style2)
    sheet1.write(x+8, 3, xlwt.Formula("SUM(D{}:D{})".format(x+4,x+8)), style2)

    x = x + 9
    return x

