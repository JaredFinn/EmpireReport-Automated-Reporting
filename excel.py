# Writing to an excel 
# sheet using Python
import time
import xlwt
from xlwt import Workbook
import xlrd
from datetime import datetime, timedelta
from datetime import date
import random
import os


adImp = 0
emailImp = 0
linkImp = 0
tweetImp = 0
adClicks = 0
emailClicks = 0
linkClicks= 0
tweetClicks = 0


def createReport(title, IMPORTNAMES, IMPORTVIEWS, IMPORTHOVERS, IMPORTCLICKS, addEmail, addLink, addTweets, videoAds, addUnique):
    global adImp
    global emailImp
    global linkImp
    global tweetImp

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

    sheet1.write(5, 0, "Banner/Video Advertisement", style)
    sheet1.write(5, 1, "Views", style2)
    sheet1.write(5, 2, "Hovers", style2)
    sheet1.write(5, 3, "Clicks", style2)

    totalViews = 0
    totalHovers = 0
    totalClicks = 0
    j = 6
    for i in IMPORTNAMES:
        sheet1.write(j, 0, i, style)
        if(("Video" in i) | ("video" in i)):
            videoAds = True
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
        adImp = x+1
        sheet1.write(x, 1, xlwt.Formula("SUM(B7:B{})".format(x)), style2)
        sheet1.write(x, 2, xlwt.Formula("SUM(C7:C{})".format(x)), style2)
        sheet1.write(x, 3, xlwt.Formula("SUM(D7:D{})".format(x)), style2)

    x = x + 1
    if(addEmail == True):
        x = addEmailToSheet(sheet1, style, style2, x, addUnique)
    if(addLink == True):
        x = addLinkToSheet(sheet1, style, style2, x)
    if(addTweets == True):
        x = addTweetsToSheet(sheet1, style, style2, x)


    if(addEmail == True or addLink == True or addTweets == True):
        sheet1.write(x+2, 0, "GRAND TOTAL:", style)
        if(addEmail == True and addLink == True and addTweets == True):
            sheet1.write(x+2, 1, xlwt.Formula("B{}+B{}+B{}+B{}".format(adImp, emailImp, linkImp, tweetImp)), style2)
            sheet1.write(x+2, 2, xlwt.Formula("C{}+C{}+C{}+C{}".format(adImp, emailImp, linkImp, tweetImp)), style2)
            sheet1.write(x+2, 3, xlwt.Formula("D{}+D{}+D{}+D{}".format(adImp, emailImp, linkImp, tweetImp)), style2)
        elif(addEmail == True and addLink == True and addTweets == False):
            sheet1.write(x+2, 1, xlwt.Formula("B{}+B{}+B{}".format(adImp, emailImp, linkImp)), style2)
            sheet1.write(x+2, 2, xlwt.Formula("C{}+C{}+C{}".format(adImp, emailImp, linkImp)), style2)
            sheet1.write(x+2, 3, xlwt.Formula("D{}+D{}+D{}".format(adImp, emailImp, linkImp)), style2)
        elif(addEmail == True and addLink == False and addTweets == True):
            sheet1.write(x+2, 1, xlwt.Formula("B{}+B{}+B{}".format(adImp, emailImp, tweetImp)), style2)
            sheet1.write(x+2, 2, xlwt.Formula("C{}+C{}+C{}".format(adImp, emailImp, tweetImp)), style2)
            sheet1.write(x+2, 3, xlwt.Formula("D{}+D{}+D{}".format(adImp, emailImp, tweetImp)), style2)
        elif(addEmail == False and addLink == True and addTweets == True):
            sheet1.write(x+2, 1, xlwt.Formula("B{}+B{}+B{}".format(adImp, linkImp, tweetImp)), style2)
            sheet1.write(x+2, 2, xlwt.Formula("C{}+C{}+C{}".format(adImp, linkImp, tweetImp)), style2)
            sheet1.write(x+2, 3, xlwt.Formula("D{}+D{}+D{}".format(adImp, linkImp, tweetImp)), style2)
        elif(addEmail == True and addLink == False and addTweets == False):
            sheet1.write(x+2, 1, xlwt.Formula("B{}+B{}".format(adImp, emailImp)), style2)
            sheet1.write(x+2, 2, xlwt.Formula("C{}+C{}".format(adImp, emailImp)), style2)
            sheet1.write(x+2, 3, xlwt.Formula("D{}+D{}".format(adImp, emailImp)), style2)
        elif(addEmail == False and addLink == True and addTweets == False):
            sheet1.write(x+2, 1, xlwt.Formula("B{}+B{}".format(adImp, linkImp)), style2)
            sheet1.write(x+2, 2, xlwt.Formula("C{}+C{}".format(adImp, linkImp)), style2)
            sheet1.write(x+2, 3, xlwt.Formula("D{}+D{}".format(adImp, linkImp)), style2)
        elif(addEmail == False and addLink == False and addTweets == True):
            sheet1.write(x+2, 1, xlwt.Formula("B{}+B{}".format(adImp, tweetImp)), style2)
            sheet1.write(x+2, 2, xlwt.Formula("C{}+C{}".format(adImp, tweetImp)), style2)
            sheet1.write(x+2, 3, xlwt.Formula("D{}+D{}".format(adImp, tweetImp)), style2)
    else:
        sheet1.write(x+1, 0, "TOTAL:", style)
        sheet1.write(x+1, 1, xlwt.Formula("SUM(B7:B{})".format(x)), style2)
        sheet1.write(x+1, 2, xlwt.Formula("SUM(C7:C{})".format(x)), style2)
        sheet1.write(x+1, 3, xlwt.Formula("SUM(D7:D{})".format(x)), style2)
    

    fileName = "{} {}.xls".format(title,date)
    filePath = "C:\Jared\EmpireReport\Reports\\Automated\\"+ fileName
    wb.save(filePath)
    os.startfile(filePath)

    totals = [totalViews, totalHovers, totalClicks]

    return totals, videoAds, filePath

def addEmailToSheet(sheet1, style, style2, x, addUnique):
    global emailImp

    sheet1.write(x+2, 0, "Email Blast w/ sponsored message", style)
    sheet1.write(x+2, 1, "Impressions", style2)
    sheet1.write(x+2, 3, "Clicks", style2)
    date = datetime.date(datetime.now())

    if(addUnique == True):
        sheet1.write(x+8, 0, "{} Unique Blast".format(date), style)
        sheet1.write(x+8, 1, 0, style2)
        sheet1.write(x+8, 3, 0, style2)
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

    if(addUnique == True):
        sheet1.write(x+9, 0, "SUBTOTAL:", style)
        emailImp = x+9+1
        sheet1.write(x+9, 1, xlwt.Formula("SUM(B{}:B{})".format(x+4,x+9)), style2)
        #sheet1.write(x+9, 2, xlwt.Formula("SUM(C{}:C{})".format(x+3,x+7)), style2)
        sheet1.write(x+9, 3, xlwt.Formula("SUM(D{}:D{})".format(x+4,x+9)), style2)
    else:
        sheet1.write(x+8, 0, "SUBTOTAL:", style)
        emailImp = x+8+1
        sheet1.write(x+8, 1, xlwt.Formula("SUM(B{}:B{})".format(x+4,x+8)), style2)
        #sheet1.write(x+9, 2, xlwt.Formula("SUM(C{}:C{})".format(x+3,x+7)), style2)
        sheet1.write(x+8, 3, xlwt.Formula("SUM(D{}:D{})".format(x+4,x+8)), style2)

    x = x+9
    return x

def addLinkToSheet(sheet1, style, style2, x):
    global linkImp
    sheet1.write(x+2, 0, "Sponsored Link", style)
    sheet1.write(x+2, 1, "Impressions", style2)
    sheet1.write(x+2, 3, "Clicks", style2)

    sheet1.write(x+3, 0, "Title of link", style)
    linkImp = x+3+1

    sheet1.write(x+3, 1, 0, style2)

    sheet1.write(x+3, 3, 0, style2)

    #sheet1.write(x+4, 0, "SUBTOTAL:", style)
    #sheet1.write(x+4, 1, xlwt.Formula("SUM(B{}:B{})".format(x+13,x+13)), style2)
    #sheet1.write(x+9, 2, xlwt.Formula("SUM(C{}:C{})".format(x+3,x+7)), style2)
    #sheet1.write(x+4, 3, xlwt.Formula("SUM(D{}:D{})".format(x+13,x+13)), style2)

    x = x+4
    return x

def addTweetsToSheet(sheet1, style, style2, x):
    global tweetImp
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
    tweetImp = x+8+1
    sheet1.write(x+8, 1, xlwt.Formula("SUM(B{}:B{})".format(x+4,x+8)), style2)
    sheet1.write(x+8, 2, xlwt.Formula("SUM(C{}:C{})".format(x+4,x+8)), style2)
    sheet1.write(x+8, 3, xlwt.Formula("SUM(D{}:D{})".format(x+4,x+8)), style2)

    x = x + 9
    return x


