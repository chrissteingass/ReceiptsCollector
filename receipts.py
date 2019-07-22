# This script logs into a Gmail account dedicated to receiving Uber ride receipts using the Gmail API. Once logged in, it will take each email in the inbox (not those in any other folders) and extract from each message the total amount charged by Uber and the date at which the ride was taken. It will also save each incoming email as an HTML file. Then, the script will compose one email for each receipt that contains the client/project, the payment method, the purpose, the date, and the total amount of that receipt. It'll attach the original receipt to the email and forward it to receipts@interworks.com and Cc: your own email address.

#####################################################################

# Update this information for each user
NAME = 'Chris'
FROM = 'steingass.iwreceipts@gmail.com'
CC = 'chrissteingass@gmail.com'

PATH = '/Users/cas/Dropbox/_Code/IW_receipts'

# Feeling unsure? Change the TO email address to something else to test out the script first
TO = 'chrissteingass@interworks.com'

# Example
# NAME = 'Chris Steingass'
# FROM = 'steingass.iwreceipts@gmail.com'
# CC = 'chris.steingass@interworks.com'
# PATH = '/Users/cas/Dropbox/_Code/IW_receipts' (MAC)
# PATH = 'C:\\uber_receipts_bouncer' (WINDOWS)

# These parts will stay the same for each outgoing email
WHAT = 'Rideshare'
WITH = 'Corporate AMEX'
SUBJECT = NAME + ' -- New Rideshare Receipt'

# The client / project will change regularly, so we'll prompt the user to enter it whenever the script is run
WHO = input('Please enter client or project')

#####################################################################

# Import packages
#
# from IPython.core.interactiveshell import InteractiveShell
# InteractiveShell.ast_node_interactivity = "all"

import os

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from googleapiclient.discovery import build

from httplib2 import Http
from oauth2client import file, client, tools
import base64

from bs4 import BeautifulSoup

import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Check for OS
def check_os():
    if os.name == 'posix':
        return('OSX')
    elif os.name == 'nt':
        return('Windows')

CURRENT_OS = check_os()
#####################################################################

# Set up connection to Gmail API
SCOPES = 'https://mail.google.com/'

store = file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
    creds = tools.run_flow(flow, store)

service = build('gmail', 'v1', http=creds.authorize(Http()))

# Call the Gmail API
all_emails = service.users().messages().list(userId='me').execute()

# Fetch IDs of all emails in account
email_ids_list = [email_id['id'] for email_id in all_emails['messages']]

# Create empty dictionary. We'll populate this with the body of the outgoing message.
emails = {}
for email_id in email_ids_list:
    emails.update({email_id : {}})

# We only want to work with the emails in the inbox, so we'll delete any email ID that is not in the inbox from our emails dictionary
for i, email_id in enumerate(email_ids_list):
    try:
        labels = service.users().messages().get(userId='me', id=email_id, format='metadata').execute()['labelIds']

        if 'INBOX' not in labels:
            del emails[email_id]
    except:
        del emails[email_id]

# For each email id, in the emails dictionary, request the entire email and navigate to the body of the text. Then, pull from that raw html code the price of the ride and the date of the ride. Then, for each email, compose an outgoing message that contains this information along with client/project, payment method, and purpose.
for email_id in emails:
    payload = service.users().messages().get(userId='me', id=email_id, format='full').execute()['payload']

    try:
        body = payload['parts'][0]['body']['data']
        body_decoded = base64.urlsafe_b64decode(body.encode('ASCII'))
        soup = BeautifulSoup(body_decoded, 'html.parser')

        date_string_raw = str(soup.find_all('span', class_ = "Uber18_text_p1 black"))
        total_string_raw = str(soup.find_all('span', class_ = 'Uber18_text_p2'))

        date_string = date_string_raw[date_string_raw.find('15px;') + 8 : -9]
        total_string = total_string_raw[total_string_raw.find('$') : -9]

        outgoing_message = 'WHAT: %s \n WHO: %s \n WITH: %s \n WHEN: %s \n TOTAL: %s \n \n \n Please find the receipt attached as an HTML file' % (WHAT, WHO, WITH, date_string, total_string)

        emails[email_id]['body'] = outgoing_message

    except:
        continue

# For each email id in our emails list, get the entire text and save it to an HTML file. We'll use this to include to our outgoing emails as an attachment.
for email_id in emails:
    raw_encoded = service.users().messages().get(userId='me', id=email_id, format='full').execute()['payload']['parts'][0]['body']['data']

    try:
        receipt_message = base64.urlsafe_b64decode(raw_encoded.encode('ASCII'))

        soup = BeautifulSoup(receipt_message, 'html.parser')

        if CURRENT_OS == 'OSX':
            path = PATH + 'email_attachments/'
        elif CURRENT_OS == 'Windows':
            path = PATH + 'email_attachments\\'

        html_file = open(path + email_id + '.htm', 'w')
        html_file.write(str(soup))
        html_file.close()

    except:
        continue


# For each email in our inbox, send a message to receipts@interworks containing a summary of the receipt in the body and the original receipt as an attachment. Once done, archive the email.
for email_id in emails:
    message = MIMEMultipart()

    message['From'] = FROM
    message['To'] = TO
    message['Cc'] = CC
    message['Subject'] = SUBJECT

    body = emails[email_id]['body']

    message.attach(MIMEText(body, 'plain'))

    if CURRENT_OS == 'OSX':
        path = PATH + '/email_attachments/'
    elif CURRENT_OS == 'Windows':
        path = PATH + '\\email_attachments\\'

    filename = email_id + '.htm'
    attachment = open(path + email_id + '.htm', 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    message.attach(part)

    message = {'raw' : base64.urlsafe_b64encode(message.as_string().encode()).decode()}

    service.users().messages().send(userId = 'me', body = message).execute()

    service.users().messages().modify(userId='me', id=email_id, body={'removeLabelIds': ['INBOX']}).execute()
