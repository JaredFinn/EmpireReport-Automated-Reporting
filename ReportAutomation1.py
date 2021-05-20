import argparse
import tkinter as tk
from tkinter import *
from tkinter.ttk import Combobox

from googleapiclient.discovery import build
import httplib2
from oauth2client import client
from oauth2client import file
from oauth2client import tools

SCOPES = ['https://www.googleapis.com/auth/analytics.readonly']
CLIENT_SECRETS_PATH = 'client_secrets.json' # Path to client_secrets.json file.
VIEW_ID = '115062057'
CLICK_VIEWS = []


root = tk.Tk()
root.title("Report Automation")

canvas = tk.Canvas(root, height=600, width= 600, bg="#4a98f0")
canvas.pack()

frame = tk.Frame(root, bg="white")
frame.place(relwidth=0.9, relheight=0.4, relx=0.05, rely=0.05)

emailFrame = tk.Frame(root, bg="white")
emailFrame.place(relwidth=0.9, relheight=0.45, relx=0.05, rely=0.5)
emailLabel=Label(emailFrame, text="Drafted Email", bg="white")
emailLabel.place(x=10, y=5)

email = Text(emailFrame, bg="grey")
email.pack(padx=40, pady=30)

data=("Current Story", "Past Story", "Ad Report")
cb=Combobox(root, values=data)
cb.place(x=250, y=100)
 
all = '/'  

reportLabel=Label(root, text="Select Report Type", bg="white")
reportLabel.place(x=135, y=100)
storyLabel= Label(root, text="Enter Story", bg="white")
storyLabel.place(x=135, y=150)

storyInput = Entry(root)
storyInput.place(x=210, y=150)


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

def get_reportPastStory(analytics, rgx):
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

def get_reportAds(analytics, rgx):
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

      for i, values in enumerate(dateRangeValues):
        print('Date range (' + str(i) + ')')
        for metricHeader, value in zip(metricHeaders, values.get('values')):
          print( metricHeader.get('name') + ': ' + value+"\n\n")
          CLICK_VIEWS.append(value)

  email.insert(END, "Recipient,\n\n")
  email.insert(END, "Your Story " + rgx + " is doing great!\n" + "\nIt generated " + str(CLICK_VIEWS[0]) + " link clicks and " + str(CLICK_VIEWS[1]) + " impressions.\nScreenshots from google analytics attached.\n\n")
  email.insert(END, "Thank You!\n")
  email.insert(END, "Best Regards,\n")
  email.insert(END, "JP")

def report():
  type=cb.get()
  rgx=storyInput.get()
  #story = Label(root, text=type)
  #storyURL = Label(root, text=rgx)
  #story.place(x=300, y=200)
  #storyURL.place(x=300, y=220)
  if(type == "Current Story"):
    analytics = initialize_analyticsreporting()
    response = get_reportCurrentStory(analytics, rgx)
    print_response(response, rgx)
  elif(type == "Past Story"):
    analytics = initialize_analyticsreporting()
    response = get_reportPastStory(analytics, rgx)
    print_response(response, rgx)
  else:
    analytics = initialize_analyticsreporting()
    response = get_reportAds(analytics, rgx)
    print_response(response, rgx)


button1 = tk.Button(text='Report', command=report)
button1.place(x=210, y=200)


root.mainloop()


