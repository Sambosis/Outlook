# email_processor.py
EXCHANGE_EMAIL='gpgaskin@autochlor.com'
EXCHANGE_DOMAIN_USERNAME='autochlor\\gpgaskin'
EXCHANGE_PASSWORD='redtag19'
EXCHANGE_SERVER='london.autochlor.net'
EXCHANGE_VERSION='Exchange2016'
OUTPUT_DIR='gpg2/'
TIMEZONE='US/Eastern'
DAYS_AGO=1

import os
import re
import logging
from datetime import datetime, timedelta
from exchangelib import Credentials, Account, DELEGATE, Message, FileAttachment, ItemAttachment, Configuration
from exchangelib.errors import ErrorTooManyObjectsOpened
import pytz
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
email = EXCHANGE_EMAIL
domain_username = EXCHANGE_DOMAIN_USERNAME
password = EXCHANGE_PASSWORD
server = EXCHANGE_SERVER
version = EXCHANGE_VERSION
output_dir = OUTPUT_DIR
timezone_name = TIMEZONE
days_ago = DAYS_AGO

def sanitize_filename(filename):
    """Sanitize the filename by removing or replacing invalid characters."""
    invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
    filename = re.sub(invalid_chars, '_', filename)
    filename = filename.rstrip('. ')
    if not filename:
        filename = 'unnamed'
    if len(filename) > 245:  
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
    recipient_name = item.to_recipients[0].name if item.to_recipients else 'Unknown_Recipient'
    subject = sanitize_filename(item.subject) if item.subject else 'No_Subject'
    email_out = f"{output_dir}to_{recipient_name} - {subject} - {item.datetime_received.strftime('%I-%M%p %m-%d%Y')}"

    # Save the email contents to a file named email.html
    email_filename = f"{email_out}.html"
    try:
        with open(email_filename, 'w', encoding='utf-8') as f:
            f.write(f"<html><body>\n")
            f.write(f"<h1>Subject: {item.subject}</h1>\n")
            f.write(f"<p><strong>Received:</strong> {item.datetime_received}</p>\n")
            f.write(f"<p><strong>Sender:</strong> {item.sender.email_address}</p>\n")
            to_addresses = ', '.join([r.email_address for r in item.to_recipients if r.email_address])
            f.write(f"<p><strong>To:</strong> {to_addresses}</p>\n")
            f.write(f"<p><strong>Body:</strong></p>\n")
            f.write(f"{item.body}\n")
            f.write(f"</body></html>\n")
    except Exception as e:
        logging.error(f"Error writing email file {email_filename}: {str(e)}")

    # Save attachments
    attachment_dir = f"{output_dir}{sanitize_filename(recipient_name)}/"
    Path(attachment_dir).mkdir(parents=True, exist_ok=True)
    for attachment in item.attachments:
        if isinstance(attachment, FileAttachment):
            safe_attachment_name = sanitize_filename(attachment.name)
            attachment_filename = f"{attachment_dir}{safe_attachment_name}"
            try:
                with open(attachment_filename, 'wb') as f:
                    f.write(attachment.content)
            except Exception as e:
                logging.error(f"Error saving attachment {attachment.name}: {str(e)}")
        elif isinstance(attachment, ItemAttachment):
            try:
                attached_item = attachment.item
                attached_subject = sanitize_filename(attached_item.subject)
                attached_email_filename = f"{attachment_dir}attached_email_{attached_subject}_{attached_item.datetime_received.strftime('%Y%m%d%H%M%S')}.html"
                with open(attached_email_filename, 'w', encoding='utf-8') as f:
                    f.write(f"<html><body>\n")
                    f.write(f"<h1>Subject: {attached_item.subject}</h1>\n")
                    f.write(f"<p><strong>Received:</strong> {attached_item.datetime_received}</p>\n")
                    f.write(f"<p><strong>Sender:</strong> {attached_item.sender.email_address}</p>\n")
                    f.write(f"<p><strong>Body:</strong></p>\n")
                    f.write(f"{attached_item.body}\n")
                    f.write(f"</body></html>\n")
            except Exception as e:
                logging.error(f"Error saving attached email: {str(e)}")

def main():
    # Load configuration from environment variables
    email = EXCHANGE_EMAIL
    domain_username = EXCHANGE_DOMAIN_USERNAME
    password = EXCHANGE_PASSWORD
    server = EXCHANGE_SERVER
    version = EXCHANGE_VERSION
    output_dir = OUTPUT_DIR
    timezone_name = TIMEZONE
    days_ago = DAYS_AGO

    # Setup exchange connection
    credentials = Credentials(username=domain_username, password=password)
    config = Configuration(server=server, credentials=credentials)
    account = Account(
        primary_smtp_address=email,
        credentials=credentials,
        autodiscover=True,
        access_type=DELEGATE,
        config=config
    )

    # Calculate time frame
    local_tz = pytz.timezone(timezone_name)
    time_frame = local_tz.localize(datetime.now() - timedelta(days=days_ago))

    # Ensure output directory exists
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Process emails
    logging.info("Processing sent emails...")
    for _ in process_email(account, account.sent, output_dir, time_frame):
        pass

    logging.info("Processing inbox emails...")
    for _ in process_email(account, account.inbox, output_dir, time_frame):
        pass

if __name__ == "__main__":
    main()