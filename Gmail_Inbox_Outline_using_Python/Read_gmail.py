"""
Problem Statement:
Using python,
- read all emails
- download the attachments
- write its contents (Date-Time, From, Subject) to an excel sheet

Before executing this file, please make sure that you've turned on 'Less secure app access' in your GMail. To do that:
Log in GMail acc -> Click on profile icon -> Manage your google acc -> Security -> Less secure app access -> ON

Prepared by: Dhairya Shah
Given by: Santosh Sir (SAC)
Date: 31 March 2022
"""

# Importing required libraries
import imaplib
import email
import traceback
import os
import pandas as pd

# Getting user information
user_email = input('Enter E-Mail ID: ')  # Your E-Mail ID
user_pwd = input('Enter Passcode: ')  # Your Passcode
folder_dir = input('Enter directory to store files: ')  # Directory where attachments and excel sheet are to be stored

# Creating working folders and files
os.chdir(folder_dir)  # Changes working directory as mentioned by the user
if 'Email_Data' not in os.listdir():  # Creates a folder in folder_dir where files will be stored
    os.mkdir('Email_Data')
os.chdir(folder_dir + '/Email_data')  # Changes working directory to the newly created folder

if 'Attachments' not in os.listdir():
    os.mkdir('Attachments')
os.chdir(os.curdir + '/Attachments')

# Logging into account
mail = imaplib.IMAP4_SSL("imap.gmail.com")  # Connects to an IMAP4 server over SSL encryption
mail.login(user_email, user_pwd)  # Logs into your GMail using the given credentials
mail.select('inbox')  # Goes into the inbox section of GMail

my_dict = {"ID": [], "Date": [], "From": [], "Subject": []}  # Empty dictionary to store items


def read_email():

    global msg
    try:
        search = mail.search(None, 'ALL')  # Searches for all the mails in the inbox
        mail_ids = search[1]  # Outputs the number of mails present in the inbox in the range (1, n)
        id_list = mail_ids[0].split()  # Assigns each mail a unique ID

        stop = int(id_list[-12])
        start = int(id_list[-1])

        for i in range(start, stop, -1):  # Iterates through the inbox
            data = mail.fetch(str(i), '(RFC822)')  # Fetches a single email from its ID using RFC822 protocol

            # Getting data to be added in Excel
            for response in data:
                arr = response[0]
                if isinstance(arr, tuple):  # The isinstance() function returns True if the specified object is
                    # of the specified type, otherwise False.
                    msg = email.message_from_string(str(arr[1], 'utf-8'))  # 'email' is a python package in which the
                    # message class provides functionality for accessing and querying header files and mail body.

                    # Appends 4 features (id, date, from, sub) of a mail to a dictionary on every iteration
                    my_dict['ID'].append(str(i))
                    my_dict['Date'].append(msg['Date'])
                    my_dict['From'].append(msg['From'])
                    my_dict['Subject'].append(msg['Subject'])

            # Downloading attachments
            for part in msg.walk():  # .walk() function is a generator which is used to iterate over all parts of the
                # message
                if part.get_content_maintype() == 'multipart':  # Checks if main content type of mail is multipart
                    continue
                if part.get('Content-Disposition') is None:  # Content-Disposition indicates that content is to be
                    # displayed either inline in the browser or as an attachment that is downloaded locally
                    continue

                file_name = part.get_filename()  # Return the value of filename parameter of Content-Disposition (inline
                # or attachment) header of the message

                if bool(file_name):
                    file_path = os.path.join(os.curdir, file_name)
                    if not os.path.isfile(file_path):
                        fp = open(file_path, 'wb')  # Opens file in write mode in binary format
                        fp.write(part.get_payload(decode=True))
                        fp.close()

    except Exception as e:
        traceback.print_exc()  # Printing stack trace for exception to get more info of error
        print(str(e))


def write_excel():
    os.chdir(folder_dir + '/Email_data')  # Current directory comes out of Attachments folder

    df = pd.DataFrame(data=my_dict)  # Converts the dictionary into a pandas dataframe
    df.to_excel('Excel_Data.xlsx')


read_email()
write_excel()
