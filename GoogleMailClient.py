""" 
https://developers.google.com/gmail/api/quickstart/python?hl=es-419

The code allows to extract information from the headers of the emails stored in the Inbox folder of the Gmail account.

After requesting user authentication (if necessary), the code retrieves all emails from the Inbox folder (up to 500) and iterates over each one of them.
For each email message, it extracts the value of the 'From' header and create a dictionary whose keys are the unique senders and the values are the number of messages sent by that sender and the sum of their respective message size estimates. Then, it stores this information in a CSV file named 'senders.csv' in a directory called 'data'. 

The following libraries must be imported beforehand: 
- os
- from googleapiclient.discovery import build
- from google.oauth2.credentials import Credentials
- from google.auth.transport.requests import Request
- from google_auth_oauthlib.flow import InstalledAppFlow
- csv

To run the script, you will need to:
- Enable the Gmail API via the Google Developers Console
- Create and download your own 'credentials.json' file and store it in the same directory as the script
- Set the SCOPES
- Run the script

Importantly, before running the script again, you need to delete the 'token.json' file. """

import base64
import csv
import os
import re
import yaml # pip install pyyaml
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

import utils

global config2

class GmailClient:
 def __init__(self,config):
      self.config=config
      self.service=self.get_gmail_service()
      
   
 #---------------------------------------------------
 # Get Gmail Service
 #---------------------------------------------------
 def get_gmail_service(self):
 # Connect to Gmail API using credentials

  # If modifying these scopes, delete the file token.json.
  SCOPES=self.config['SCOPES']
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists(self.config['token.json']):
   creds = Credentials.from_authorized_user_file(self.config['token.json'], SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
       creds.refresh(Request())
    else:
       flow = InstalledAppFlow.from_client_secrets_file(
       self.config['credentials.json'], SCOPES)
       creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open(self.config['token.json'], 'w') as token:
     token.write(creds.to_json())

  service = build("gmail", "v1", credentials=creds)
  return service

#----------------------------------------------------------------
# Write a file with list of senders, number of emails, total size
#----------------------------------------------------------------
 def all_senders(self):

    nextPageToken = ''
    messages = []
    
    #Messages list
    while True:
        # All messages in INBOX
        query=""
        if self.config['start_date']!="": 
           start_date=self.config['start_date']
           query+=f'after:{start_date} '
        if self.config['end_date']!="":
           end_date=self.config['end_date']
           query+=f'before:{end_date} '

        results = self.service.users().messages().list(
           userId='me', 
           pageToken=nextPageToken,
           maxResults=500,
           labelIds=self.config['labels_ids'],
           q=query
        ).execute()
    
        # We add messages to complete list
        messages.extend(results['messages'])
        print(f"Number of messages : {len(messages)}",end='\r')
        
        #We check if there is another page
        if 'nextPageToken' in results:
            nextPageToken = results['nextPageToken']
            #break # To test code
        else:
            break

    #For each message we retrieve information
    senders = {}
    i=0
    print()
    for message in messages:        
        msg = self.service.users().messages().get(userId='me', id=message['id'], format='metadata').execute()
        i+=1
        print(f"Message : {i}",end='\r')
        headers = msg['payload']['headers']
        header = list(filter(lambda headers: (headers['name'] == 'From' or headers['name'] == 'from'), headers))        
        if len(header)!=0:
            sender = header[0]['value']
            if sender in senders.keys(): 
                senders[sender]['count']+=1 
                senders[sender]['size']+=msg['sizeEstimate']
            else: 
                senders[sender]={}
                senders[sender]['count']= 1
                senders[sender]['size']=msg['sizeEstimate']

    
    # We store information in file
    current_datetime = datetime.now()
    date_str = current_datetime.strftime("%Y%m%d_%H%M%S")
    filename=self.config['output_dir']+'/senders_'+date_str+'.csv'
    with open(filename, mode='w', newline='',encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['From','Count','Size'])
        for sender in sorted(senders.keys()):
            writer.writerow([sender,senders[sender]['count'],senders[sender]['size']])

    print(f'Info saved in {filename}')

#---------------------------------------------------
# Define a function to save email messages as .eml files
 def save_email_to_disk(self,msg_id, headers,directory):
    try:
        # Retrieve the email message
        msg = self.service.users().messages().get(userId="me", id=msg_id, format='raw').execute()

        # Decode the message content from base64url
        msg_part = msg['raw']
        msg_bytes = base64.urlsafe_b64decode(msg_part.encode('UTF-8'))

        for header in headers:
           if header['name'] == 'Date':
            email_date = header['value']
           elif header['name'] == 'Subject':
            subject = header['value']        

        date_obj = datetime.strptime(email_date, "%a, %d %b %Y %H:%M:%S %z")
        timestamp = int(date_obj.timestamp())  
        string_date=date_obj.strftime('%Y%m%d_%H%M%S')
        subject_ok=utils.sanitize_name(subject)
        filepath = os.path.join(directory, f'{string_date} - {subject_ok}.eml')

        # Write the message content to a .eml file
        with open(filepath, 'wb') as f:
            f.write(msg_bytes)
        # Upate the modified date to the original one
        os.utime(filepath, (timestamp, timestamp))
    except HttpError as error:
        print(f'An error occurred: {error}')
        msg_bytes = None
    return msg_bytes

#---------------------------------------------------
# Retrieve all messages from specific sender
#---------------------------------------------------
 def get_all_messages_from_sender(self, sender):
    # Define query for filtering emails
    query = f"from:{sender} "
    if self.config['start_date']!="": 
      start_date=self.config['start_date']
      query+=f'after:{start_date} '
    if self.config['end_date']!="":
      end_date=self.config['end_date']
      query+=f'before:{end_date} '    

    # Use the API to get all emails matching the query
    results = self.service.users().messages().list(
       userId="me",
       labelIds=self.config['labels_ids'],
       q=query
    ).execute()

    # Get the email ID for each matching email
    emails = [r["id"] for r in results.get("messages", [])]

    return emails

#---------------------------------------------------
# print all messages from specific sender
#---------------------------------------------------
 def print_all_messages_from_sender(self, sender):

    emails=self.get_all_messages_from_sender(sender)

    # Get the full email data for each email
    print(f"Emails de {sender}")
    for email_id in emails:
        # Retrieve the email message
        msg = self.service.users().messages().get(userId="me", id=email_id).execute()
        headers = msg['payload']['headers']
        for header in headers:
           if header['name'] == 'Date':
            date = header['value']
           elif header['name'] == 'Subject':
            subject = header['value']
        print(date + " - "+ subject) 

#---------------------------------------------------
# save all messages from specific sender
#---------------------------------------------------
 def save_all_messages_from_sender(self, sender):
    emails=self.get_all_messages_from_sender(sender)

    # Get the full email data for each email
    print(f"\nEmails from {sender}")
    dir_sender=self.config['output_dir']+"/"+utils.sanitize_name(sender)
    if(not os.path.exists(dir_sender)): os.mkdir(dir_sender)    
    
    i=1
    for email_id in emails:
      msg = self.service.users().messages().get(userId="me", id=email_id).execute()
      headers = msg['payload']['headers']  
      i+=1    
      print(f"Saving message : {i}/{len(emails)}",end='\r')
      self.save_email_to_disk(email_id,headers,dir_sender)    
    
#---------------------------------------------------
# delete all messages from specific sender
#---------------------------------------------------
 def delete_all_messages_from_sender(self, sender):
    emails=self.get_all_messages_from_sender(sender)   
    
    i=0
    for email_id in emails:
      self.service.users().messages().delete(userId="me", id=email_id).execute()
      i+=1    
      print(f"Deleting message : {i}/{len(emails)}",end='\r') 

#---------------------------------------------------
# menu print all messages from specific sender
#---------------------------------------------------
def menu_print_messages(my_gmail):
   ar_senders=utils.inputML("Sender/s")
   if len(ar_senders)!=0:
      for sender in ar_senders:
         print(f"Messages from : {sender}...")
         my_gmail.print_all_messages_from_sender(sender)

#---------------------------------------------------
# menu save all messages from specific sender
#---------------------------------------------------
def menu_save_messages(my_gmail):
   ar_senders=utils.inputML("Sender/s")
   if len(ar_senders)!=0:
      for sender in ar_senders:
         print(f"Saving messages from : {sender}...")
         my_gmail.save_all_messages_from_sender(sender)
   
#---------------------------------------------------
# menu delete all messages from specific sender
#---------------------------------------------------
def menu_delete_messages(my_gmail):
   ar_senders=utils.inputML("Sender/s")
   if len(ar_senders)!=0:
      for sender in ar_senders:
         print(f"Deleting messages from : {sender}...")
         my_gmail.delete_all_messages_from_sender(sender)

#---------------------------------------------------
# Set filters
#---------------------------------------------------
def set_filters():
   print("Variable to set (start_date,end_date,labels): ",end='')
   l_var=input()
   if l_var not in ('start_date','end_date','labels'): return
   print(f'{l_var}: ',end='')
   l_value=input()
   if l_value=='': return
   if l_var=='labels':
    try:
     temp_list = eval(l_value)
     if not isinstance(temp_list, list):
        print(f"{l_value} does not represent a list!")
     else:
        for element in temp_list:
            if not isinstance(element, str):
                print(f"{element} is not a string")
                break
        else:
            config2['labels_ids']=l_value
    except Exception as e:
     print("Error: ", e)
   else:
    try:
      datetime.strptime(l_value, '%Y/%m/%d')
      config2[l_var]=l_value
    except ValueError:
      print("Invalid date format")      


#---------------------------------------------------
# Main process
#---------------------------------------------------
# Read config file
with open('./config.yml', 'r') as file:
    # Load the YAML data into a Python object
    config2 = yaml.load(file, Loader=yaml.FullLoader)

# We validate date format
if config2['start_date']!="":
    date_obj = datetime.strptime(config2['start_date'], "%Y/%m/%d")
if config2['end_date']!="":
    date_obj = datetime.strptime(config2['end_date'], "%Y/%m/%d")

if (not os.path.exists(config2['output_dir'])): os.mkdir(config2['output_dir'])


my_gmail = GmailClient(config2)

l_answer=""
print()
while l_answer!="0":
 print()
 print("-----------------------------")
 print("GMAIL BASIC CLIENT")
 print("-----------------------------")
 print(f"start_date: {config2['start_date']}")
 print(f"end_date: {config2['end_date']}")
 print(f"labels: {config2['labels_ids']}")
 print()
 print("1 . Set filters")
 print("2 . Retrieve senders to disk")
 print("3 . List emails from sender/s")
 print("4 . Save emails from sender/s to disk")
 print("5 . Delete emails from sender/s")
 print("0 . Exit")
 print("Your option : ",end='')
 l_answer=input()
 if l_answer=="1": set_filters()
 if l_answer=="2": my_gmail.all_senders()
 if l_answer=="3": menu_print_messages(my_gmail)
 if l_answer=="4": menu_save_messages(my_gmail)
 if l_answer=="5": menu_delete_messages(my_gmail)
  