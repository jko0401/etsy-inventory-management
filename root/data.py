from pathlib import Path
from gmail_functions import *
import pandas as pd
import base64

# get the Gmail API service
service = gmailAuthenticate()

transaction_emails = searchMessages(service, 'from: transaction@etsy.com')
email_ids = [temp['id'] for temp in transaction_emails]

SALES_DICT = {}
TRANS_DICT = {}
SHIP_DICT = {}

for eid in email_ids:
    msg_items = readMessage(service, eid)
    for header in msg_items[0]:
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
        email_body = base64.urlsafe_b64decode(msg_items[1][0]['body']['data']).decode("utf-8")
        populateSalesInfo(orderno, email_body, SALES_DICT)
        populateTransInfo(orderno, email_body, TRANS_DICT)

shipping_emails = searchMessages(service, 'from:no-reply@etsy.com')
shipemail_ids = [temp['id'] for temp in shipping_emails]


for sid in shipemail_ids:
    msg_items = readMessage(service, sid)
    email_body = base64.urlsafe_b64decode(msg_items[1][1]['body']['data']).decode("utf-8")
    try:
        l_id = populateShipInfo(email_body, SHIP_DICT)

        for header in msg_items[0]:
            name = header.get("name")
            value = header.get("value")
            if name == "Subject":
                orderno = value.split("#", 1)[1]
                SHIP_DICT[l_id]['order_number'] = orderno
            if name == "Date":
                created_date = value
                SHIP_DICT[l_id]['date'] = created_date
    except IndexError:
        print(sid, 'Bundled Labels')


mypath = Path().absolute()
filePath = mypath/'statements'
df = pd.DataFrame()
for file in filePath:
    if file.endswith('.csv'):
        df = df.append(pd.read_csv(file), ignore_index=True)

t_df = df[df['Type'] == 'Transaction']
s_df = df[df['Type'] == 'Sale']

missing_ship = set()
missing_trans = set()

for index, row in t_df.iterrows():
    if row['Title'] == 'Shipping':
        for key, value in SHIP_DICT.items():
            if value['order_number'] == row['Info'][7:]:
                SHIP_DICT[key][value['trans_cost']] = row['Fees & Taxes']
            else:
                missing_ship.add(row['Info'][7:])
    else:
        try:
            TRANS_DICT[row['Info'][13:]]['cost'] = row['Fees & Taxes']
        except KeyError:
            missing_trans.add(row['Info'][13:])

missing_order = set()
for index, row in s_df.iterrows():
    try:
        SALES_DICT[row['Title'][-10:]]['tax_fee'] = row['Fees & Taxes']
    except KeyError:
        missing_order.add(row['Title'][-10:])

print('Missing Transactions:', missing_trans)
print('Missing Shippings:', missing_ship)
print('Missing Orders:', missing_order)

trans_df = pd.DataFrame(TRANS_DICT).transpose().reset_index().rename(columns={'index': 'transaction_id'})
sales_df = pd.DataFrame(SALES_DICT).transpose().reset_index().rename(columns={'index': 'order_number'})
ship_df = pd.DataFrame(SHIP_DICT).transpose().reset_index().rename(columns={'index': 'label_id'})




