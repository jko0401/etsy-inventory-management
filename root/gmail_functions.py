import os
import re
import pickle
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = 'test@gmail.com'

def gmailAuthenticate():
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


def searchMessages(service, query):
    result = service.users().messages().list(userId='me', q=query).execute()
    messages = [ ]
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me', q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages


def readMessage(service, message_id):
    msg = service.users().messages().get(userId='me', id=message_id).execute()
    payload = msg['payload']
    headers = payload.get("headers")
    parts = payload.get("parts")
    return headers, parts


def findPart(part, body):
    result = re.findall(part, body)
    try:
        for i, x in enumerate(result):
            result[i] = x.strip(' ').strip('\r')
    except AttributeError:
        pass
    return result


def populateSalesInfo(ordernumber, body, sales_dict):
    info_list = [('Order Total:(.*)', 'order_total'), ('Shipping:(.*)', 'customer_shipping'), ('Discount(.*)', 'discount')]
    for info in info_list:
        sales_dict[ordernumber][info[1]] = findPart(info[0], body)


def populateTransInfo(ordernumber, body, trans_dict):
    trans_ids = findPart('Transaction ID:(.*)', body)
    for t_id in trans_ids:
        trans_dict[t_id] = {'order_number': ordernumber, 'cost': ''}


def populateShipInfo(body, ship_dict):
    label = findPart('(\d{12})(?=</a>)', body)
    label_id = label[0]
    ship_dict[label_id] = {'date': '', 'order_number': '', 'cost': '', 'address': '', 'trans_cost': '',
                           'transaction_id': ''}
    info_list = [('(?=\$)(.*)(?= USD)', 'cost'), ('(?<!Ships To)(<span (.*)</span>)', 'address'),
                 ('(?<=Transaction ID: )(\d{10})', 'transaction_id')]
    for info in info_list:
        if info[1] is 'cost':
            ship_dict[label_id][info[1]] = findPart(info[0], body)[-1]
        elif info[1] is 'address':
            ship_dict[label_id][info[1]] = cleanHtml(findPart(info[0], body)[0][0])
        else:
            ship_dict[label_id][info[1]] = findPart(info[0], body)[0]

    return label_id


def cleanHtml(raw_html):
    cleanr = re.compile('<.*?>|&([a-z0-9]+|#[0-9]{1,6}|#x[0-9a-f]{1,6});')
    cleantext = re.sub(cleanr, ' ', raw_html).strip()
    return cleantext
