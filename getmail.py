"""
Fetches emails from the user's Exchange account inbox and sent items, and saves the emails and attachments to the local file system.

The script uses the exchangelib library to connect to the Exchange server and retrieve emails. It filters the emails based on a time frame (10 days in this case) and saves the email contents and attachments to the local file system.

For each email, the script creates a directory named with the recipient's name and the email subject, and saves the email contents to an HTML file. It also saves any attachments to a subdirectory named with the recipient's name.

If the email has an attached email item, the script saves the contents of the attached email item to a separate HTML file.
"""
# The `os` module provides a way of interacting with the operating system. It is likely used elsewhere in the file to perform file system operations, such as creating directories.
#Imports the os module, which provides a way of interacting with the operating system. This is likely used elsewhere in the file to perform file system operations, such as creating directories.
import os
import re
from exchangelib import Credentials, Account, DELEGATE, Message, FileAttachment, ItemAttachment, Configuration
from datetime import datetime, timedelta
import pytz

# Replace with your actual email, domain, username, and password
email = 'gpgaskin@autochlor.com'
domain_username = 'autochlor\\gpgaskin'  # Use DOMAIN\username format
password = 'redtag19'
server = 'london.autochlor.net'
version = 'Exchange2016'  # Adjust based on your actual Exchange version

# Set up credentials with the 
# domain
credentials = Credentials(username=domain_username, password=password)

# Manually configure the server settings
config = Configuration(server=server, credentials=credentials)
account = Account(
    primary_smtp_address = email,
    credentials = credentials,
    autodiscover = True,
    access_type = DELEGATE
)

# Calculate the date 30 days ago with timezone awareness
local_tz = pytz.timezone('US/Eastern')  # Replace with your timezone, e.g., 'UTC', 'US/Eastern', etc.
sent_time_frame = local_tz.localize(datetime.now() - timedelta(days=2))

local_tz = pytz.timezone('US/Eastern') 

from_time_frame= local_tz.localize(datetime.now() - timedelta(days=2))

# Ensure the output directory exists
output_dir = f'gpg2/'
os.makedirs(output_dir, exist_ok=True)

def sanitize_filename(filename):
    """
    Sanitize the filename by removing or replacing invalid characters.
    """
    # Replace invalid characters with underscore
    # This covers Windows and Unix illegal characters
    invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
    filename = re.sub(invalid_chars, '_', filename)
    
    # Remove trailing spaces and periods (problematic on Windows)
    filename = filename.rstrip('. ')
    
    # Ensure filename isn't empty and isn't too long (Windows has 255 char limit)
    if not filename:
        filename = 'unnamed'
    if len(filename) > 245:  # leaving room for extension
        filename = filename[:245]
    
    return filename

def process_email(account, email_folder, output_dir, time_frame):
    """Process emails in the specified folder."""
    try:
        for item in email_folder.filter(datetime_received__gte=time_frame).order_by('-datetime_received'):
            if isinstance(item, Message):
                yield process_email_item(account, item, output_dir)
    except ErrorTooManyObjectsOpened as e:
        logging.error(f"Too many objects error: {e}")

def process_email_item(account, item, output_dir):
    """Process a single email item."""
    recipient_name = sanitize_filename(item.to_recipients[0].name) if item.to_recipients else 'Unknown_Recipient'
    subject = sanitize_filename(item.subject) if item.subject else 'No_Subject'
    received_time = item.datetime_received.strftime('%m-%d-%Y_%I-%M%p')
    email_out = f"to_{recipient_name} - {subject} - {received_time}"
    
    # Save the email contents to a file named email.html
    email_filename = f"{output_dir}{email_out}.html"
    try:
        with open(email_filename, 'w', encoding='utf-8') as f:
            f.write(f"<html><body>\n")
            f.write(f"<h1>Subject: {item.subject}</h1>\n")
            f.write(f"<p><strong>Received:</strong> {item.datetime_received}</p>\n")
            f.write(f"<p><strong>Sender:</strong> {item.sender.email_address}</p>\n")
            to_addresses = ', '.join([recipient.email_address for recipient in item.to_recipients if recipient.email_address])
            f.write(f"<p><strong>To:</strong> {to_addresses}</p>\n")
            f.write(f"<p><strong>Body:</strong></p>\n")
            f.write(f"{item.body}\n")
            f.write(f"</body></html>\n")
    except Exception as e:
        print(f"Error writing email file {email_filename}: {str(e)}")
    
    print(f"Email saved to {email_out}.html")
    attachment_dir = f"{output_dir}{email_out}_attachments/"
    # Create the attachment directory
    os.makedirs(attachment_dir, exist_ok=True)
    
    # Save attachments
    for attachment in item.attachments:
        if isinstance(attachment, FileAttachment):
            try:
                safe_attachment_name = sanitize_filename(attachment.name)
                attachment_filename = f"{attachment_dir}{safe_attachment_name}"
                with open(attachment_filename, 'wb') as f:
                    f.write(attachment.content)
                print(f"Saved attachment: {attachment_filename}")
            except Exception as e:
                print(f"Error saving attachment {attachment.name}: {str(e)}")
        elif isinstance(attachment, ItemAttachment):
            try:
                attached_item = attachment.item
                attached_subject = sanitize_filename(attached_item.subject)
                attached_received_time = attached_item.datetime_received.strftime('%m-%d-%Y_%I-%M%p')
                attached_email_filename = f"{attachment_dir}attached_{attached_subject}_{attached_received_time}.html"
                with open(attached_email_filename, 'w', encoding='utf-8') as f:
                    f.write(f"<html><body>\n")
                    f.write(f"<h1>Subject: {attached_item.subject}</h1>\n")
                    f.write(f"<p><strong>Received:</strong> {attached_item.datetime_received}</p>\n")
                    f.write(f"<p><strong>Sender:</strong> {attached_item.sender.email_address}</p>\n")
                    f.write(f"<p><strong>Body:</strong></p>\n")
                    f.write(f"{attached_item.body}\n")
                    f.write(f"</body></html>\n")
                print(f"Saved attached email: {attached_email_filename}")
            except Exception as e:
                print(f"Error saving attached email: {str(e)}")

# Fetch sent emails from the last 30 days
for _ in process_email(account, account.sent, output_dir, sent_time_frame):
    pass

output_dir = f'gpg2/'
os.makedirs(output_dir, exist_ok=True)
i = 0
# Fetch emails from the last 30 days
for item in account.inbox.filter(datetime_received__gte=from_time_frame).order_by('-datetime_received'):
    # print(f"now on item {i}")
    i += 1
    if isinstance(item, Message):
        # Extract sender email address
        
        from_address = item.sender.email_address.replace(' ', '_').replace('/', '_')
        from_name = item.sender.name.replace(' ', '_').replace('/', '_')
        # Create a directory for each email based on sender and subject
        if item.subject:
            subject = item.subject.replace(' ', '_').replace('/', '_')
        else:
            subject = 'No_Subject'
        email_filename = f"{output_dir}from_{from_name}_{subject}_{item.datetime_received.strftime('%Y%m%d%H%M%S')}"

        # Save the email contents to a file named email.html
        email_filename = f"{email_filename}.html"
        try:
            with open(email_filename, 'w', encoding='utf-8') as f:
                f.write(f"<html><body>\n")
                f.write(f"<h1>Subject: {item.subject}</h1>\n")
                f.write(f"<p><strong>Received:</strong> {item.datetime_received}</p>\n")
                f.write(f"<p><strong>Sender:</strong> {item.sender.email_address}</p>\n")
                f.write(f"<p><strong>Body:</strong></p>\n")
                f.write(f"{item.body}\n")
                f.write(f"</body></html>\n")
        except Exception as e:
            print(f"Error writing email file {email_filename}: {str(e)}")

        email_out = f"{from_address} - {subject} - {item.datetime_received.strftime('%I:%M%p %m-%d-%y')}"
        print(f"Email  to {email_filename}")
        attachment_dir = f"{output_dir}{from_name}/"
        # Save attachments
        for attachment in item.attachments:
            os.makedirs(attachment_dir, exist_ok=True)

            if isinstance(attachment, FileAttachment):
                try:
                    safe_attachment_name = sanitize_filename(attachment.name)
                    attachment_filename = f"{attachment_dir}{safe_attachment_name}"
                    with open(attachment_filename, 'wb') as f:
                        f.write(attachment.content)
                except Exception as e:
                    print(f"Error saving attachment {attachment.name}: {str(e)}")

            elif isinstance(attachment, ItemAttachment):
                # If the attachment is an email item, save its contents similarly
                try:
                    attached_item = attachment.item
                    attached_subject = sanitize_filename(attached_item.subject)
                    attached_email_filename = f"{output_dir}attached_{attached_subject}_{attached_item.datetime_received.strftime('%Y%m%d%H%M%S')}.html"
                    with open(attached_email_filename, 'w', encoding='utf-8') as f:
                        f.write(f"<html><body>\n")
                        f.write(f"<h1>Subject: {attached_item.subject}</h1>\n")
                        f.write(f"<p><strong>Received:</strong> {attached_item.datetime_received}</p>\n")
                        f.write(f"<p><strong>Sender:</strong> {attached_item.sender.email_address}</p>\n")
                        f.write(f"<p><strong>Body:</strong></p>\n")
                        f.write(f"{attached_item.body}\n")
                        f.write(f"</body></html>\n")
                except Exception as e:
                    print(f"Error saving attached email: {str(e)}")

                email_out = f"{from_address} - {item.datetime_received.strftime('%I:%M%p %m-%d-%y')}"
                print(f"Attach to {attached_email_filename}")
#  - {attached_item.name}