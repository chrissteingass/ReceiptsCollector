# This script logs into a Gmail account dedicated to receiving Uber ride receipts using the Gmail API. Once logged in, it will take each email in the inbox (not those in any other folders) and extract from each message the total amount charged by Uber and the date at which the ride was taken. It will also save each incoming email as an HTML file. Then, the script will compose one email for each receipt that contains the client/project, the payment method, the purpose, the date, and the total amount of that receipt. It'll attach the original receipt to the email and forward it to receipts@interworks.com and Cc: your own email address.

#####################################################################

PATH = '/Users/cas/Dropbox/__projects/IW_receipts/v3'

WHAT = 'Rideshare'
WITH = 'Corporate AMEX'

#####################################################################

# Import packages

from IPython.core.interactiveshell import InteractiveShell
InteractiveShell.ast_node_interactivity = "all"

import os
import datetime
import time

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from googleapiclient.discovery import build

from __future__ import print_function
from httplib2 import Http
from oauth2client import file, client, tools
import base64

from bs4 import BeautifulSoup

import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

os.chdir('/Users/cas/Dropbox/__projects/IW_receipts/v3/')
#####################################################################

# Set up connection to Gmail API
def connect():
    SCOPES = 'https://mail.google.com/'

    store = file.Storage(PATH + '/token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)

    service = build('gmail', 'v1', http=creds.authorize(Http()))

    return(service)

service = connect()

#####################################################################
# Helpers

# Update labels list
def update_labels():
    all_labels_raw = service.users().labels().list(userId='me').execute()['labels']
    all_labels = pd.DataFrame({
        'label_name' : [x['name'] for x in all_labels_raw],
        'label_id' : [x['id'] for x in all_labels_raw]
        })
    return(all_labels)

# Generate list of all email ids in the inbox folder
def generate_inbox_email_ids():
    # Call the Gmail API
    all_emails = service.users().messages().list(userId='me').execute()

    # Fetch IDs of all emails in account
    email_ids_list = [email_id['id'] for email_id in all_emails['messages']]

    # Filter down email_ids_list to only those emails in the inbox
    inbox_email_ids =[]
    for email_id in email_ids_list:
        try:
            labels = service.users().messages().get(userId='me', id=email_id, format='minimal').execute()['labelIds']

            if 'INBOX' in labels:
                inbox_email_ids.append(email_id)
        except:

            continue

    return(inbox_email_ids)

# Generate list of all email ids in the alias inbox folder
def generate_alias_inbox_email_ids(alias):
    # Call the Gmail API
    all_emails = service.users().messages().list(userId='me').execute()

    # Fetch list of label ids
    all_labels = update_labels()

    # Fetch IDs of all emails in account
    email_ids_list = [email_id['id'] for email_id in all_emails['messages']]

    # Filter down email_ids_list to only those emails in the folder
    alias_inbox_email_ids =[]
    for email_id in email_ids_list:
        try:
            labels = service.users().messages().get(userId='me', id=email_id, format='minimal').execute()['labelIds']

            label_id = all_labels.label_id[all_labels.label_name == alias + '_rideshare_inbox'].values[0]

            if label_id in labels:
                alias_inbox_email_ids.append(email_id)
        except:
            continue

    return(alias_inbox_email_ids)

# Generate list of all new request emails
def generate_request_email_ids():
    # Call the Gmail API
    all_emails = service.users().messages().list(userId='me').execute()

    # Fetch list of label ids
    all_labels = update_labels()

    # Fetch IDs of all emails in account
    email_ids_list = [email_id['id'] for email_id in all_emails['messages']]

    # Filter down email_ids_list to only those emails in the
    request_email_ids =[]
    for email_id in email_ids_list:
        try:

            labels = service.users().messages().get(userId='me', id=email_id, format='minimal').execute()['labelIds']

            label_id = all_labels.label_id[all_labels.label_name == '_request_new'].values[0]

            if label_id in labels:
                request_email_ids.append(email_id)

        except:

            continue

    return(request_email_ids)

#####################################################################

# Sort mail into subfolders
def sort_into_alias_inbox():

    # Generate list of emails in inbox folder
    inbox_email_ids = generate_inbox_email_ids()

    new_mail = len(inbox_email_ids)

    if new_mail == 0:
        print('No emails in inbox')
    elif new_mail > 1:
        print(str(new_mail) + ' emails in inbox')
    elif new_mail == 1:
        print(str(new_mail) + ' email in inbox')

    # Remove emails from inbox and place them in alias inbox
    if len(inbox_email_ids) != 0:
        print('Sorting receipts into alias folders')

        #email_id = '1662b38aa8f0402c'
        for email_id in inbox_email_ids:
            try:
                sender = str(service.users().messages().get(userId='me', id=email_id, format='metadata').execute()['payload']['headers'][0]['value'])

                if sender.find('+') != -1:
                    alias = sender[sender.find('+') + 1 : sender.find('@')]
                    alias_inbox = alias + '_rideshare_inbox'
                    alias_sent = alias + '_rideshare_sent'

                    while True:
                        all_labels = update_labels()
                        if alias_inbox in all_labels.label_name.values:
                            label_id = all_labels.label_id[all_labels.label_name == alias_inbox].values[0]

                            silent = service.users().messages().modify(userId='me', id=email_id, body={'removeLabelIds': ['INBOX'], 'addLabelIds' : [label_id]}).execute()

                        else:
                            label = {
                                'messageListVisibility': 'show',
                                'labelListVisibility': 'labelShow'}

                            label['name'] = alias_inbox
                            silent = service.users().labels().create(userId='me', body=label).execute()

                            label['name'] = alias_sent
                            silent = service.users().labels().create(userId='me', body=label).execute()

                            continue

                        print('--- Receipt sorted into ' + alias_inbox)
                        break
            except:
                ('--- Skipped non-receipt email')
                continue

# Execute requests: Send receipt emails
def execute_new_requests():
    # Fetch list of new requests
    request_email_ids = generate_request_email_ids()
    if len(request_email_ids) != 0:
        print('New request found')
    else:
        print('No new requests')
    # Fetch list of labels
    all_labels = update_labels()

    # For each request email, fetch client code and alias from request email and send message to receipts@interworks.com
    for request_email_id in request_email_ids:
        print('Executing new request')
        form_input_raw = service.users().messages().get(userId='me', id=request_email_id, format='metadata').execute()['snippet']

        client_code = form_input_raw[form_input_raw.find('Client Code') + len('CLient Code') + 1: form_input_raw.find('BounceMyReceipts Alias')-1]

        print('--- Client code: ', client_code)

        alias = form_input_raw[form_input_raw.find('BounceMyReceipts Alias') + len('BounceMyReceipts alias') + 1 : form_input_raw.find('***')].lower()

        print('--- Alias: ', alias)

        # Fetch list of email ids in alias inbox folder
        email_ids_list = generate_alias_inbox_email_ids(alias)

        # Create empty dictionary. We'll populate this with the body of the outgoing message.
        emails = {}
        for email_id in email_ids_list:
            emails.update({email_id : {}})

        # For each email id, in the emails dictionary, request the entire email and navigate to the body of the text. Then, pull from that raw html code the price of the ride and the date of the ride. Then, for each email, compose an outgoing message that contains this information along with client/project, payment method, and purpose.

        # SEND MESSAGES TO RECEIPTS@
        print('Generating outgoing messages')
        for i, email_id in enumerate(emails):
            payload = service.users().messages().get(userId='me', id=email_id, format='full').execute()['payload']

            if payload['headers'][24]['value'] == 'Lyft Ride Receipt <no-reply@lyftmail.com>':
                try:
                    body = payload['body']['data']
                    body_decoded = base64.urlsafe_b64decode(body.encode('ASCII'))
                    soup = BeautifulSoup(body_decoded, 'html.parser')

                    date_string_raw = str(soup.find_all('span', class_ = 'dt-transaction'))
                    total_string_raw = str(soup.find_all('strong'))

                    date_string = date_string_raw[30:-8]
                    total_string = total_string_raw[total_string_raw.find('$'):-36]

                    outgoing_message = 'FROM: %s \n WHAT: %s \n WHO: %s \n WITH: %s \n WHEN: %s \n TOTAL: %s \n \n \n Please find the receipt attached as an HTML file' % (str(alias).replace('.', ' ').title(), WHAT, client_code, WITH, date_string, total_string)

                    print('\n')
                    print('Generated new outgoing message:')
                    print(outgoing_message)
                    print('\n')

                    emails[email_id]['body'] = outgoing_message

                    path = PATH + '/_email_attachments/'

                    html_file = open(path + email_id + '.htm', 'w')
                    html_file.write(str(soup))
                    html_file.close()

                except:
                    print('email', str(i), email_id, 'failed')
                    continue

            else:
                try:
                    body = payload['parts'][0]['body']['data']
                    body_decoded = base64.urlsafe_b64decode(body.encode('ASCII'))
                    soup = BeautifulSoup(body_decoded, 'html.parser')

                    date_string_raw = str(soup.find_all('span', class_ = "Uber18_text_p1 black"))
                    total_string_raw = str(soup.find_all('span', class_ = 'Uber18_text_p2'))

                    date_string = date_string_raw[date_string_raw.find('15px;') + 8 : -9]
                    total_string = total_string_raw[total_string_raw.find('$') : -9]

                    outgoing_message = 'FROM: %s \n WHAT: %s \n WHO: %s \n WITH: %s \n WHEN: %s \n TOTAL: %s \n \n \n Please find the receipt attached as an HTML file' % (str(alias).replace('.', ' ').title(), WHAT, client_code, WITH, date_string, total_string)

                    print('\n')
                    print('Generated new outgoing message:')
                    print(outgoing_message)
                    print('\n')

                    emails[email_id]['body'] = outgoing_message

                    path = PATH + '/_email_attachments/'

                    html_file = open(path + email_id + '.htm', 'w')
                    html_file.write(str(soup))
                    html_file.close()

                except:
                        print('email', str(i), email_id, 'failed')
                        continue

        # For each email in our inbox, send a message to receipts@interworks containing a summary of the receipt in the body and the original receipt as an attachment. Once done, archive the email.
        print('Sending messages to receipts@interworks.com')
        for email_id in emails:
            time.sleep(1)
            message = MIMEMultipart()

            message['From'] = 'BounceMyReceipts@gmail.com'
            message['To'] = 'receipts@interworks.com'
            message['Cc'] = alias + '@interworks.com'
            message['Subject'] = str(alias).replace('.', ' ').title() + ' --- new rideshare receipt'

            body = emails[email_id]['body']

            message.attach(MIMEText(body, 'plain'))

            path = PATH + '/_email_attachments/'

            filename = email_id + '.htm'
            attachment = open(path + email_id + '.htm', 'rb')

            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

            message.attach(part)

            message = {'raw' : base64.urlsafe_b64encode(message.as_string().encode()).decode()}

            silent = service.users().messages().send(userId = 'me', body = message).execute()

            alias_LabelId_inbox = all_labels.label_id[all_labels.label_name == alias + '_rideshare_inbox'].values[0]
            alias_LabelId_sent = all_labels.label_id[all_labels.label_name == alias + '_rideshare_sent'].values[0]

            time.sleep(1)

            silent = service.users().messages().modify(userId='me', id=email_id, body={'removeLabelIds': ['INBOX', alias_LabelId_inbox], 'addLabelIds' : [alias_LabelId_sent]}).execute()

            print('--- Message sent')

        request_new_LabelId = all_labels.label_id[all_labels.label_name == '_request_new'].values[0]

        request_old_LabelId = all_labels.label_id[all_labels.label_name == '_request_old'].values[0]

        silent = service.users().messages().modify(userId='me', id=request_email_id, body={'removeLabelIds': [request_new_LabelId], 'addLabelIds' : [request_old_LabelId]}).execute()

        print('Archived request')

# Run
def execute():
    sort_into_alias_inbox()
    execute_new_requests()

# Trigger
while True:
    try:
        print('\n', datetime.datetime.now(), '\n')
        execute()
        print('\n', '...')
        time.sleep(600)
    except:
        print('Error. Trying again in 10 minutes.')
        time.sleep(600)
        continue




execute()
