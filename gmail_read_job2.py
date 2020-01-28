#!/usr/bin/env python
# coding: utf-8

# In[1]:


from __future__ import print_function
import pickle
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
import pandas as pd
import datetime
from datetime import date
from datetime import datetime
import re
import sys


# In[2]:


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.

JobStart = datetime.now()
print("Start time for job is {}".format(JobStart))

print(os.getcwd())
if os.path.exists('token_sample.pickle'):
    with open('token_sample.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials_sample.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token_sample.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('gmail', 'v1', credentials=creds)

list_of_ids = []
content_of_msgs = []

file_path = "Your directory path"
file_name = "sender_details.xlsx"
file_dtls = file_path+file_name
output_file_name = "read_email_dets"

if os.path.exists(file_dtls):
    print("file exists")
    pdf_file_dtls = pd.read_excel(file_dtls)
else:
    print("file not there...please provide file")
    raise SystemExit("Exiting as program can not move forward without file")
print("location of file is {}".format(file_dtls))

sender = pdf_file_dtls["Retailer"].iloc[0]

sender_dtls = pd.read_excel(file_path + output_file_name + "_" + sender + ".xlsx",index=False)

if sender_dtls[sender_dtls["Status"] == 'pending'].empty:
    raise sys.exit("Exiting as all the emails for sender {} are read".format(sender))
else:
    start_date = datetime.strftime(sender_dtls[sender_dtls["Status"] == 'pending'][0:1]["StartDate"].iloc[0],'%Y/%m/%d')
    end_date = datetime.strftime(sender_dtls[sender_dtls["Status"] == 'pending'][0:1]["EndDate"].iloc[0],'%Y/%m/%d')
    sender_email = sender_dtls[sender_dtls["Status"] == 'pending'][0:1]["Email"].iloc[0]


data_file = file_path + sender + ".csv"
id_file = file_path + sender + "_ids.csv"

query = ("'" + 'from:({}) after:{} before:{}'.format(sender_email,start_date, end_date) + "'")
print(query)
# Call the Gmail API
#results = service.users().getProfile(userId='me').execute()
results = service.users().messages().list(userId='me',q=query).execute()
#print("first run number of email %s " % results['resultSizeEstimate'])
messages = []
if 'messages' in results:
    messages.extend(results['messages'])
while 'nextPageToken' in results:
    page_token = results['nextPageToken']
    results = service.users().messages().list(userId='me',q=query,
                                              pageToken=page_token).execute()
    messages.extend(results['messages'])
    print("page token %s  " % page_token)
for i in range(len(messages)):
    list_of_ids.append(messages[i]['id'])
    try:

        message = service.users().messages().get(userId='me', id=messages[i]['id'], format='metadata',
                                                 metadataHeaders=['Subject', 'Date', 'To', 'From','snippet']).execute()

        message_dict = message['payload']['headers']
        snippet_dict = ({u'name':'snippet',u'value':message['snippet']})
        message_dict.append(snippet_dict)

        content_of_msgs.append(message_dict)


    except errors.HttpError:
        print('An error occurred: %s' % errors.HttpError)
            
print("Number of Messages Read {}".format(len(content_of_msgs)))
          
df2 = pd.DataFrame()
for i in range(len(content_of_msgs)):
    df1 = pd.DataFrame(content_of_msgs[i]).sort_values(by=['name']).transpose()
    df1 = df1.drop(['name'])
    df1.columns = ['Date','From','Subject','To','Snippet']
    df2 = df2.append(df1,ignore_index=True)

df2['Subject2'] = df2.Subject.str.encode(encoding="ascii", errors="ignore")
df2 = df2.drop(columns=['Subject'])
df2 = df2.rename(columns={'Subject2':'Subject'})

df2['Snippet2'] = df2.Snippet.str.encode(encoding="ascii", errors="ignore")
df2 = df2.drop(columns=['Snippet'])
df2 = df2.rename(columns={'Snippet2':'Snippet'})

if not os.path.exists(data_file):
    print("file does not exist...Creating file")
    df2.to_csv(data_file,index=False)
else:
    print("file already exist...appending file")
    with open(data_file, 'ab') as d:
        df2.to_csv(d,header=False,index=False)
if not os.path.exists(id_file):
    print("id file does not exist...Creating file")
    pd.DataFrame(list_of_ids).to_csv(id_file,header=False,index=False)
    print("id file created")
else:
    print("id file already exist...appending file")
    with open(id_file, 'ab') as d:
        pd.DataFrame(list_of_ids).to_csv(d,header=False,index=False)
    print("id file appended")

sender_dtls["Status"][sender_dtls["StartDate"] == datetime.strptime(start_date,'%Y/%m/%d')] = 'Done'
sender_dtls["NoofEmails"][sender_dtls["StartDate"] == datetime.strptime(start_date,'%Y/%m/%d')] = len(content_of_msgs)
sender_dtls["JobStart"][sender_dtls["StartDate"] == datetime.strptime(start_date,'%Y/%m/%d')] = JobStart
sender_dtls["JobEnd"][sender_dtls["StartDate"] == datetime.strptime(start_date,'%Y/%m/%d')] = datetime.now()

sender_dtls.to_excel(file_path + output_file_name + "_" + sender + ".xlsx",index=False)
print("End time for job is {}".format(datetime.now()))


# In[ ]:




