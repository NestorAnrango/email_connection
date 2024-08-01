# -*- coding: utf-8 -*-
"""
Created on Thu Jun  6 16:24:54 2024

@author: nestor.gualsaqui

Description:
This code is a collection from Stackoverflow. The goal of this code is to create a connection with Outlook
and automatically send an email with an attach to a specific person or group.

Code structure:
1. Email connection
2. Find the newest
3. Send email, with all information needed

"""

import win32com.client as win32
import os


def send_email_with_attachment(to, subject, body, attachment_path):
    # Create a new Outlook application session
    outlook = win32.Dispatch('outlook.application')
    
    # Create a new email item
    mail = outlook.CreateItem(0)
    
    # Set email parameters
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    
    # Check if the attachment exists
    if os.path.isfile(attachment_path):
        mail.Attachments.Add(attachment_path)
    else:
        print(f"Attachment not found: {attachment_path}")
        return
    
    # Send the email
    mail.Send()
    print("Email sent successfully...")




# -----------------------------------------------------------------------------
def get_newest_file(directory):
    """
    This funtion get the newest file in a folder
    :param directory: String path of the data to find e.g 'C:\path\to\data'
    :return: the newest file name in a folder e.g Newest_file.xlsx
    """
    print('Getting newest file...')
    # Get a list of all files in the specified directory
    files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]

    # Check if there are any files
    if not files:
        print("No files found.")
        return None
    # Sort the files by modification time (newest first)
    files.sort(key=lambda f: os.path.getmtime(os.path.join(directory, f)), reverse=True)
    # Get the newest file
    # newest_file = os.path.join(directory, files[0])
    latest_file = files[0]
    print(f'Newest file: {latest_file}...')
    return latest_file


# -------------------------------------------------------------------------------
# Example usage
cwd = os.getcwd()
latest_report = get_newest_file(directory)
full_path_report = os.path.join(cwd, 'data', 'results', latest_report)
to = 'your_email@gmail.com'
subject = 'Montly report'
body = 'Hi, this is an automatically generated email.'
attachment_path = full_path_report

# sending email:
send_email_with_attachment(to, subject, body, attachment_path)
