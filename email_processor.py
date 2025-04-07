# email_processor.py
EXCHANGE_EMAIL='gpgaskin@autochlor.com'
EXCHANGE_DOMAIN_USERNAME='autochlor\\gpgaskin'
EXCHANGE_PASSWORD='redtag19'
EXCHANGE_SERVER='london.autochlor.net'
EXCHANGE_VERSION='Exchange2016'
OUTPUT_DIR='gpg2/'
TIMEZONE='US/Eastern'
DAYS_AGO=5

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

def replace_cid_urls(body, email_out, output_dir):
    """
    Replace 'cid:' URLs in the email body with valid HTTP URLs pointing to attachment files.
    """
    if body is None:
        return ""
        
    try:
        def cid_replacer(match):
            cid = match.group(1)
            # Assuming the cid corresponds to the attachment filename
            attachment_url = f"/gpg2/{email_out}_attachments/{cid}"
            return f'src="{attachment_url}"'
        
        # Replace all src="cid:filename" with src="/gpg2/.../filename"
        pattern = r'src=["\']cid:(.*?)["\']'
        replaced_body = re.sub(pattern, cid_replacer, body, flags=re.IGNORECASE)
        return replaced_body
    except Exception as e:
        logging.error(f"Error replacing CID URLs: {str(e)}")
        return body if body else ""

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
    try:
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
                to_addresses = ', '.join([r.email_address for r in item.to_recipients if r.email_address])
                f.write(f"<p><strong>To:</strong> {to_addresses}</p>\n")
                
                # Replace 'cid:' URLs in the email body
                body_content = item.body if item.body else ""
                sanitized_body = replace_cid_urls(body_content, email_out, output_dir)
                f.write(f"<p><strong>Body:</strong></p>\n")
                f.write(f"{sanitized_body}\n")
                f.write(f"</body></html>\n")
        except Exception as e:
            logging.error(f"Error writing email file {email_filename}: {str(e)}")
        
        # Save attachments
        attachment_dir = f"{output_dir}{email_out}_attachments/"
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
                except Exception as e:
                    logging.error(f"Error saving attached email: {str(e)}")
        
        return email_out
    except Exception as e:
        logging.error(f"Error processing email: {str(e)}")
        return f"error_{datetime.now().strftime('%Y%m%d%H%M%S')}"

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
