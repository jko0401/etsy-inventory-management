import os
import re
import pickle
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from mimetypes import guess_type as guess_mime_type

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

def read_message(service, message_id):
    """
    This function takes Gmail API `service` and the given `message_id` and does the following:
        - Downloads the content of the email
        - Prints email basic information (To, From, Subject & Date) and plain/text parts
        - Creates a folder for each email based on the subject
        - Downloads text/html content (if available) and saves it under the folder created as index.html
        - Downloads any file that is attached to the email and saves it in the folder created
    """
    msg = service.users().messages().get(userId='me', id=message_id).execute()
    # parts can be the message body, or attachments
    payload = msg['payload']
    headers = payload.get("headers")
    parts = payload.get("parts")
    for header in headers:
        name = header.get("name")
        value = header.get("value")
        if name == "Subject":
            try:
                print("Order Number:", value.split("#",1)[1][:-1])
            except IndexError:
                if 'Etsy Order confirmation for:' in value:
                    print("Order Number:", value.split("(",1)[1][:-1])
        if name == "Date":
            # we print the date when the message was sent
            print("Date:", value)
    email_body = base64.urlsafe_b64decode(parts[0]['body']['data']).decode("utf-8")
    print('Transaction IDs: ', find_part('Transaction ID:(.*)', email_body))
    print('Items: ', find_part('Item:(.*)', email_body))
    print('Primary: ', find_part('Color:(.*)', email_body))
    print('Secondary: ', find_part('Secondary color:(.*)', email_body))
    print('Quantity: ', find_part('Quantity:(.*)', email_body))
    print('Price: ', find_part('Item price:(.*)', email_body))
    print("="*50)

def find_part(part, body):
    result = re.findall(part, body)
    try:
        for i, x in enumerate(result):
            result[i] = x.strip(' ').strip('\r')
    except AttributeError:
        pass
    return result