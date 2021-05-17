"""Hello Analytics Reporting API V4."""

import argparse

from googleapiclient.discovery import build
import httplib2
from oauth2client import client
from oauth2client import file
from oauth2client import tools

SCOPES = ['https://www.googleapis.com/auth/analytics.readonly']
CLIENT_SECRETS_PATH = 'client_secrets.json' # Path to client_secrets.json file.
VIEW_ID = '115062057'
CLICK_VIEWS = []
rgx=input("Enter Story: ")
all='/'


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

def get_report(analytics):
  # Use the Analytics Service Object to query the Analytics Reporting API V4.

  

  return analytics.reports().batchGet(
      body={
        'reportRequests': [
        {
          'viewId': VIEW_ID,
          'dateRanges': [{'startDate': '2021-05-16', 'endDate': '2021-05-16'}],
          'dimensions': [{'name': 'ga:pagePath'}],
          'metrics': [{'expression': 'ga:pageviews'}],
          'filtersExpression':f'ga:pagePath=={rgx}',
        },
        {
          'viewId': VIEW_ID,
          'dateRanges': [{'startDate': '2021-05-16', 'endDate': '2021-05-16'}],
          'dimensions': [{'name': 'ga:pagePath'}],
          'metrics': [{'expression': 'ga:pageviews'}],
          'filtersExpression':f'ga:pagePath=={all}',
        }]
      }
  ).execute()

  

def print_response(response):
  """Parses and prints the Analytics Reporting API V4 response"""
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

  print("EMAIL PREVIEW:\n")

  print("Recipient,\n")
  print("Your Story " + rgx + " is doing great!\n" + "It generated " + str(CLICK_VIEWS[0]) + " link clicks and " + str(CLICK_VIEWS[1]) + " impressions.\nScreenshots from google analytics attached.")
  print("Thank You!\n")
  print("Best Regards,")
  print("JP")


def main():

  analytics = initialize_analyticsreporting()
  response = get_report(analytics)
  print_response(response)

if __name__ == '__main__':
  main()
