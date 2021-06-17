import os
import re
import pickle
import pandas as pd
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64


# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = 'test@gmail.com'

def gmail_authenticate():
    fileDir = os.path.dirname(os.path.realpath('__file__'))
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            cred_file = os.path.join(fileDir, 'keys\credentials.json')
            flow = InstalledAppFlow.from_client_secrets_file(cred_file, SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

# get the Gmail API service
service = gmail_authenticate()

def search_messages(service, query):
    result = service.users().messages().list(userId='me',q=query).execute()
    messages = [ ]
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me',q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages

transaction_emails = search_messages(service, 'from: transaction@etsy.com')

SALES_DICT = {}


def readMessage(service, message_id):
    """

    """
    msg = service.users().messages().get(userId='me', id=message_id).execute()
    payload = msg['payload']
    headers = payload.get("headers")
    parts = payload.get("parts")
    for header in headers:
        name = header.get("name")
        value = header.get("value")
        if name == "Subject":
            if 'Purchase' in value:
                orderno = 0
                pass
            else:
                try:
                    orderno = value.split("#", 1)[1][:-1]
                    # print("Order Number:", value.split("#",1)[1][:-1])
                except IndexError:
                    orderno = value.split("(", 1)[1][:-1]
                    # print("Order Number:", value.split("(",1)[1][:-1])
        if name == "Date":
            pur_date = value
            # print("Date:", value)

    if orderno is not 0:
        SALES_DICT[orderno] = {'date': pur_date, 'order_total': '', 'tax_fee': '', 'listing_fee': '',
                               'customer_shipping': '', 'discount': '', 'profit': '', 'margin': ''}
        email_body = base64.urlsafe_b64decode(parts[0]['body']['data']).decode("utf-8")
        populateSalesInfo(orderno, email_body)


def findPart(part, body):
    result = re.findall(part, body)
    try:
        for i,x in enumerate(result):
            result[i] = x.strip(' ').strip('\r')
    except AttributeError:
        pass
    return result


def populateSalesInfo(ordernumber, body):
    info_list = [('Order Total:(.*)', 'order_total'), ('Shipping:(.*)', 'customer_shipping'), ('Discount(.*)', 'discount')]
    for info in info_list:
        SALES_DICT[ordernumber][info[1]] = findPart(info[0], body)


email_ids = [test['id'] for test in transaction_emails]
for eid in email_ids:
    readMessage(service, eid)