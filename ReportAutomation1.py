import argparse
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.ttk import Combobox
from tkinter import *
from TkinterDnD2 import *
import csv

import excel as excelTab

from googleapiclient.discovery import build
import httplib2
from oauth2client import client
from oauth2client import file
from oauth2client import tools


SCOPES = ['https://www.googleapis.com/auth/analytics.readonly']
CLIENT_SECRETS_PATH = 'client_secrets.json' # Path to client_secrets.json file.
VIEW_ID = '115062057'
CLICK_VIEWS = []
PAGES = []
IMPORTNAMES = []
IMPORTVIEWS = []
IMPORTHOVERS = []
IMPORTCLICKS = []
ADPROGRAMS = ['A More Just NYC Kivvit', 'AARP', 'Adiply', 'AFL-CIO', 'Aid in Dying'
                  ,'AMERICAN CHEMISTRY COUNCIL', 'American Legal Finance Association', 'American Progressive Plastic Bag Alliance'
                  ,'API', 'ASPCA', 'Astorino', 'Avangrid - Arch Street Communications', 'Back to Bowling'
                  ,'BASK', 'BE FAIR TO DIRECT CARE', 'BERLINROSEN ANTI-FRAUD LAWS', 'Bet on NY'
                  ,'BLUE COLLAR COALITION', 'BP AMERICAS', 'Building & Construction Trades of Greater New York'
                  ,'Bull Moose Club', 'Business Council of Westchester', 'Butler Associates'
                  ,'Cats Round Table', 'CENTRO CROMINAL JUSTICE', 'CENTRO Taxpayers for Affordable New York'
                  ,'Charter Spectrum', 'child victims act GREENBERG', 'CITIZENS FOR PROGRESS', 'Claudia Tenney for Congress'
                  ,'Clean Fuels NY Kivvit', 'Coalition for the Homeless', 'Coalition to Help Families (JACK BONNER)'
                  ,'Common Cause NY', 'Community Pharmacy Association of NYS', 'COMPASSION & CHOICES'
                  ,'Congressional Candidate', 'Cruelty Free International', 'CUNY', 'CUOMO FOR GOVERNOR'
                  ,'CWA - BERLINROSEN', 'Dev Site Test By Saad', "DON'T BLOCK NY BUILDING", 'Education Equity Campaign'
                  ,'Elise Stefanik', 'EMPIRE CITY CASINO', 'Empire Report', 'Erica Video', 'Farm Bureau'
                  ,'Friends of the BQX', 'Frontier', 'FWD.us', 'GNYHA', 'Google Adwords', 'GSG Congestion Pricing'
                  ,'GSG Criminal Justice Reform', 'GSG NYABA', 'HANYS', 'House', 'HTC ADAMS IE', 'IPPNY'
                  ,'JUUL Labs', 'KWATRA', 'Linnea Empire Test', 'Long Island Association', 'Manhattan Chamber of Commerce'
                  ,'MARATHON', 'MASK UP AMERICA', 'Metropolitan Public Strategies', 'MOLINARO', 'MPAA'
                  ,'NEW YORK YANKEES', 'New Yorkers for Clean Water and Jobs', 'New Yorkers for Responsible Gaming', 'New Yorkers United for Justice'
                  ,'NY GAMING ASSOCIATION', 'NY HEALTH ACT', 'NY League of Conservation Voters', 'NY State Industries for the Disabled'
                  ,'NY STATE WEAR A MASK CAMPAIGN', 'NYC CHARTER SCHOOLS', 'NYS Health Foundation GSG', 'NYSANA Nurse Anesthetists'
                  ,'NYSCOP L POLITI', 'NYSCOPBA', 'NYSPSP', 'NYSUT', 'NYTHA', 'Ostroff Associates', 'PARTNERSHIP FOR NYC'
                  ,'PARTNERSHIP FOR SAFE MEDICINE', 'Patrick B. Jenkins & Associates', 'PEF', 'PHRMA', 'Project Guardianship'
                  ,'psacentral.org', 'QUEENS Chamber of Commerce', 'REALITIES OF SINGLE PAYER', 'REBNY', 'Rebuild NY Now'
                  ,'Rechler Kivvit', 'Reclaim NY', 'Retail Council', 'SANDS Kivvit', "Saratoga Harness Horseperson's Assocation"
                  ,'Saratoga Mentoring', 'SEIU', 'Shenker Russo Clark', 'SIEMENS', "Sizmek's services ads", 'SKD FLEXIBLE WORK'
                  ,'SKD- RESORTS WORLD CASINO', 'SMART APPROACHES TO MARIJUANA', 'Strong Leadership NYC - Eric Adams', 'SUNY Empire State College'
                  ,'The Airbnb Tax', 'The Brooklyn Hospital Center', 'TRANSPORT WORKERS UNION', 'Trucking Association of New York'
                  ,'TRUTH ABOUT ORSTED', 'United Way', 'United Way Greater Capital Region', 'VALCOUR WIND ENERGY', 'VINCENZO GARDINO', 'WAMC', 'WESTERN OTB BATAVIA DOWNS']


root = TkinterDnD.Tk()
root.title("Report Automation")

canvas = tk.Canvas(root, height=600, width= 600, bg="#4a98f0")
canvas.pack()

frame = tk.Frame(root, bg="white")
frame.place(relwidth=0.9, relheight=0.4, relx=0.05, rely=0.05)

tabControl = ttk.Notebook(frame)

tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)

tabControl.add(tab1, text ='Email Report')
tabControl.add(tab2, text ='Excel Report')
tabControl.grid(padx=50)
tabControl.pack(expand = 1, fill ="both")


emailFrame = tk.Frame(root, bg="white")
emailFrame.place(relwidth=0.9, relheight=0.45, relx=0.05, rely=0.5)
emailLabel=Label(emailFrame, text="Drafted Email", bg="white")
emailLabel.place(x=10, y=5)

email = Text(emailFrame, bg="lightgrey")
email.pack(padx=40, pady=30)

data=("Current Story", "Past Story", "Ad Report")
ads=ADPROGRAMS

cb=Combobox(tab1, values=data)
cb.place(x=250, y=100)
 
all = '/'  

#Email Tabs
reportLabel=Label(tab1, text="Select Report Type", bg="white")
reportLabel.place(x=135, y=100)
storyLabel= Label(tab1, text="Enter Story", bg="white")
storyLabel.place(x=135, y=150)

##Excel Tab
AdvertiserLabel=Label(tab2, text="Program", bg="white")
AdvertiserLabel.place(x=40, y=100)
nameInput = Combobox(tab2, values=ads)
nameInput.place(x=100, y=100)


emailVar = BooleanVar()
emailCheck = Checkbutton(tab2, text="Email", variable=emailVar)
emailCheck.place(x=255, y=100)
linkVar = BooleanVar()
linkCheck = Checkbutton(tab2, text="Sponsored Link", variable=linkVar)
linkCheck.place(x=320, y=100)
tweetVar = BooleanVar()
tweetCheck = Checkbutton(tab2, text="Tweets", variable=tweetVar)
tweetCheck.place(x=435, y=100)

def drop(event):
        entry_sv.set(event.data)

dir = tk.Label(tab2, text="Drop CSV file from Broadstreet Ads, select program name and press report.\nFill out any remaining stats, save file, and press construct email.")
dir.pack(pady=10)

entry_sv = StringVar()
entry_sv.set('Drop Here...')
entry = Entry(tab2, textvar=entry_sv, width=80)
entry.pack(fill=X, padx=40, pady=0)
entry.drop_target_register(DND_FILES)
entry.dnd_bind('<<Drop>>', drop)


def initialize_analyticsreporting():
  """Initializes the analyticsreporting service object.

  Returns:
    analytics an authorized analyticsreporting service object.
  """
  # Parse command-line arguments.
  parser = argparse.ArgumentParser(
      formatter_class=argparse.RawDescriptionHelpFormatter,
      parents=[tools.argparser])
  flags = parser.parse_args([])

  # Set up a Flow object to be used if we need to authenticate.
  flow = client.flow_from_clientsecrets(
      CLIENT_SECRETS_PATH, scope=SCOPES,
      message=tools.message_if_missing(CLIENT_SECRETS_PATH))

  # Prepare credentials, and authorize HTTP object with them.
  # If the credentials don't exist or are invalid run through the native client
  # flow. The Storage object will ensure that if successful the good
  # credentials will get written back to a file.
  storage = file.Storage('analyticsreporting.dat')
  credentials = storage.get()
  if credentials is None or credentials.invalid:
    credentials = tools.run_flow(flow, storage, flags)
  http = credentials.authorize(http=httplib2.Http())

  # Build the service object.
  analytics = build('analyticsreporting', 'v4', http=http)

  return analytics

analytics = initialize_analyticsreporting()

def get_reportCurrentStory(analytics, rgx):
  # Use the Analytics Service Object to query the Analytics Reporting API V4.
  return analytics.reports().batchGet(
      body={
        'reportRequests': [
        {
          'viewId': VIEW_ID,
          'dateRanges': [{'startDate': '2021-05-17', 'endDate': '2021-05-17'}],
          'dimensions': [{'name': 'ga:pagePath'}],
          'metrics': [{'expression': 'ga:pageviews'}],
          'filtersExpression':f'ga:pagePath=={rgx}',
        },
        {
          'viewId': VIEW_ID,
          'dateRanges': [{'startDate': '2021-05-17', 'endDate': '2021-05-17'}],
          'dimensions': [{'name': 'ga:pagePath'}],
          'metrics': [{'expression': 'ga:pageviews'}],
          'filtersExpression':f'ga:pagePath=={all}',
        }]
      }
  ).execute()



def getPages(analytics):
  # Use the Analytics Service Object to query the Analytics Reporting API V4.
  return analytics.reports().batchGet(
      body={
        'reportRequests': [
        {
          'viewId': VIEW_ID,
          'dateRanges': [{'startDate': '2021-04-21', 'endDate': '2021-05-20'}],
          'dimensions': [{'name': 'ga:pageTitle'}],
          'metrics': [{'expression': 'ga:pageviews'}],
        },
        ]
      }
  ).execute()

def getPath(analytics):
  # Use the Analytics Service Object to query the Analytics Reporting API V4.
  return analytics.reports().batchGet(
      body={
        'reportRequests': [
        {
          'viewId': VIEW_ID,
          'dateRanges': [{'startDate': '2021-04-21', 'endDate': '2021-05-20'}],
          'dimensions': [{'name': 'ga:pageTitle'}],
          'metrics': [{'expression': 'ga:pageviews'}],
        },
        ]
      }
  ).execute()


def print_response(response, rgx):
  """Parses and prints the Analytics Reporting API V4 response"""
  email.delete(1.0, END)
  print()
  print("API DATA: \n")
  for report in response.get('reports', []):
    columnHeader = report.get('columnHeader', {})
    dimensionHeaders = columnHeader.get('dimensions', [])
    metricHeaders = columnHeader.get('metricHeader', {}).get('metricHeaderEntries', [])
    rows = report.get('data', {}).get('rows', [])

    for row in rows:
      dimensions = row.get('dimensions', [])
      dateRangeValues = row.get('metrics', [])

      for header, dimension in zip(dimensionHeaders, dimensions):
        print(header + ': ' + dimension)
        PAGES.append(dimension)

      for i, values in enumerate(dateRangeValues):
        print('Date range (' + str(i) + ')')
        for metricHeader, value in zip(metricHeaders, values.get('values')):
          print( metricHeader.get('name') + ': ' + value+"\n\n")
          CLICK_VIEWS.append(value)

  email.insert(END, "Recipient,\n\n")
  email.insert(END, "Your Story " + rgx + " is doing great!\n" + "\nIt generated a total of " + str(CLICK_VIEWS[0]) + " link clicks and " + str(CLICK_VIEWS[1]) + " impressions.\nScreenshots from google analytics attached.\n\n")
  email.insert(END, "Thank You!\n")
  email.insert(END, "Best Regards,\n")
  email.insert(END, "JP")

def print_titles(response):
  """Parses and prints the Analytics Reporting API V4 response"""
  email.delete(1.0, END)
  print()
  print("API DATA: \n")
  for report in response.get('reports', []):
    columnHeader = report.get('columnHeader', {})
    dimensionHeaders = columnHeader.get('dimensions', [])
    rows = report.get('data', {}).get('rows', [])

    for row in rows:
      dimensions = row.get('dimensions', [])

      for header, dimension in zip(dimensionHeaders, dimensions):
        print(header + ': ' + dimension)
        PAGES.append(dimension)


def report():
  type=cb.get()
  rgx=pageCb.get()
  if(type == "Current Story"):
    analytics = initialize_analyticsreporting()
    response = get_reportCurrentStory(analytics, rgx)
    print_response(response, rgx)
  #elif(type == "Past Story"):
    #analytics = initialize_analyticsreporting()
    #response = get_reportPastStory(analytics, rgx)
    #print_response(response, rgx)
  #else:
    #analytics = initialize_analyticsreporting()
    #response = get_reportAds(analytics, rgx)
    #print_response(response, rgx)


##method to create excel file and save it in location specified in excel.py
##creates total values to use for email draft
def excelReport():
  IMPORTNAMES.clear()
  IMPORTVIEWS.clear()
  IMPORTHOVERS.clear()
  IMPORTCLICKS.clear()
  #entry_sv.set(entry_sv.get()[0:-1])
  #entry_sv.set(entry_sv.get()[1:])
  importData()
  addEmail = emailVar.get()
  addLink = linkVar.get()
  addTweets = tweetVar.get()
  email.delete(1.0, END)
  title = nameInput.get()
  videoAds = False
  totals, videoAds = excelTab.createReport(title, IMPORTNAMES, IMPORTVIEWS, IMPORTHOVERS, IMPORTCLICKS, addEmail, addLink, addTweets, videoAds)
  totalsFromatted1 = '{:,.0f}'.format(totals[0])
  totalsFromatted2 = '{:,.0f}'.format(totals[1])
  totalsFromatted3 = '{:,.0f}'.format(totals[2])

  email.insert(END, "Recipient,\n\n")
  email.insert(END, "I hope that you are well!\n")
  email.insert(END, "I wanted to give you an update on the most recent video ad campaign for "+ title + ":\n\n")
  if(videoAds == True):
      email.insert(END, "Thus-far the video ads have generated " + totalsFromatted1 + " impressions, " + totalsFromatted2 + " hovers, and " + totalsFromatted3 + " link clicks.\n")
      videoAds = False
  else:
      email.insert(END, "Thus-far the banner ads have generated " + totalsFromatted1 + " impressions, " + totalsFromatted2 + " hovers, and " + totalsFromatted3 + " link clicks.\n")
      videoAds = False

  if((addEmail == True) & (addLink == False) & (addTweets == False)):
      email.insert(END, "The sponsored message in the daily email newsletter has generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "TOTAL: " + str(0) + " impressions and " + str(0) + " link clicks\n")
  elif((addEmail == False) & (addLink == True) & (addTweets == False)):
      email.insert(END, "The sponsored story on Empire Report has generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "TOTAL: " + str(0) + " impressions and " + str(0) + " link clicks\n")
  elif((addEmail == False) & (addLink == False) & (addTweets == True)):
      email.insert(END, "The sponsored tweets on Empire Report's page have generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "TOTAL: " + str(0) + " impressions and " + str(0) + " link clicks\n")
  elif((addEmail == True) & (addLink == True) & (addTweets == False)):
      email.insert(END, "The sponsored message in the daily email newsletter has generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "The sponsored story on Empire Report has generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "TOTAL: " + str(0) + " impressions and " + str(0) + " link clicks\n")
  elif((addEmail == False) & (addLink == True) & (addTweets == True)):
      email.insert(END, "The sponsored story on Empire Report has generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "The sponsored tweets on Empire Report's page have generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "TOTAL: " + str(0) + " impressions and " + str(0) + " link clicks\n")
  elif((addEmail == True) & (addLink == False) & (addTweets == True)):
      email.insert(END, "The sponsored message in the daily email newsletter has generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "The sponsored tweets on Empire Report's page have generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "TOTAL: " + str(0) + " impressions and " + str(0) + " link clicks\n")
  elif((addEmail == True) & (addLink == True) & (addTweets == True)):
      email.insert(END, "The sponsored message in the daily email newsletter has generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "The sponsored story on Empire Report has generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "The sponsored tweets on Empire Report's page have generated " + str(0) + " impressions and " + str(0) + " link clicks.\n")
      email.insert(END, "TOTAL: " + str(0) + " impressions and " + str(0) + " link clicks\n")

  email.insert(END, "Full data report is attached.\n\n")
  email.insert(END, "Thank you for working with me on this project!!\n\n")
  email.insert(END, "Best Regards,\n")
  email.insert(END, "JP Miller\n")
  email.insert(END, "Empire Report\n")
  email.insert(END, "917-565-3378")

def importData():
  with open(entry_sv.get(), 'r') as file:
    reader = csv.reader(file)
    m = 0
    for row in reader:
      if m == 0:
        m = m+1
        continue
      if row[1] != str(0):
          IMPORTNAMES.append(row[0])
          IMPORTVIEWS.append(row[1])
          IMPORTHOVERS.append(row[2])
          IMPORTCLICKS.append(row[3])


response = getPages(analytics)
print_titles(response)

pageCb=Combobox(tab1, values=PAGES)
pageCb.place(x=210, y=150)

emailBtn = tk.Button(tab1, text='Report', command=report)
emailBtn.place(x=375, y=150)

excelBtn = tk.Button(tab2, text='Report', command=excelReport)
excelBtn.place(x=175, y=175)

constructBtn = tk.Button(tab2, text='Construct Email')
constructBtn.place(x=275, y=175)

root.mainloop()