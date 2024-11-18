import os
import re
import logging
from flask import Flask, render_template, request, jsonify, send_from_directory, abort
from datetime import datetime, timedelta
from exchangelib import Credentials, Account, DELEGATE, Message, FileAttachment, ItemAttachment, Configuration
from exchangelib.errors import ErrorTooManyObjectsOpened
import pytz
from pathlib import Path
from icecream import ic 
# At the start of your Flask app
import os
EXCHANGE_EMAIL='gpgaskin@autochlor.com'
EXCHANGE_DOMAIN_USERNAME='autochlor\\gpgaskin'
EXCHANGE_PASSWORD='redtag19'
EXCHANGE_SERVER='london.autochlor.net'
EXCHANGE_VERSION='Exchange2016'
OUTPUT_DIR='gpg2/'
TIMEZONE='US/Eastern'
DAYS_AGO=1
# Ensure the email directory exists and is readable
if not os.path.exists('gpg2'):
    os.makedirs('gpg2')
app = Flask(__name__, static_folder='static')
# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Directory where emails are stored
EMAIL_DIR = '/app/gpg2'  # Update this path as needed

def sanitize_filename(filename):
    """Sanitize the filename by removing or replacing invalid characters."""
    print()
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
    print()
    try:
        for item in email_folder.filter(datetime_received__gte=time_frame).order_by('-datetime_received'):
            if isinstance(item, Message):
                yield process_email_item(account, item, output_dir)
    except ErrorTooManyObjectsOpened as e:
        logging.error(f"Too many objects error: {e}")

def process_email_item(account, item, output_dir):
    """Process a single email item."""
    print()
    recipient_name = item.to_recipients[0].name if item.to_recipients else 'Unknown_Recipient'
    subject = sanitize_filename(item.subject) if item.subject else 'No_Subject'
    email_out = f"to_{recipient_name} - {subject} - {item.datetime_received.strftime('%I-%M%p %m-%d%Y')}"
    
    # Path relative to EMAIL_DIR
    email_filename = os.path.join(output_dir, f"{email_out}.html")
    
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
    attachment_dir = os.path.join(output_dir, sanitize_filename(recipient_name))
    Path(attachment_dir).mkdir(parents=True, exist_ok=True)
    for attachment in item.attachments:
        if isinstance(attachment, FileAttachment):
            safe_attachment_name = sanitize_filename(attachment.name)
            attachment_filename = os.path.join(attachment_dir, safe_attachment_name)
            try:
                with open(attachment_filename, 'wb') as f:
                    f.write(attachment.content)
            except Exception as e:
                logging.error(f"Error saving attachment {attachment.name}: {str(e)}")
        elif isinstance(attachment, ItemAttachment):
            try:
                attached_item = attachment.item
                attached_subject = sanitize_filename(attached_item.subject)
                attached_email_filename = os.path.join(attachment_dir, f"attached_email_{attached_subject}_{attached_item.datetime_received.strftime('%Y%m%d%H%M%S')}.html")
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
    
    return email_filename  # Return the path for reference

def setup_exchange_connection():
    """Setup Exchange connection using environment variables."""
    email = os.getenv('EXCHANGE_EMAIL')
    domain_username = os.getenv('EXCHANGE_DOMAIN_USERNAME')
    password = os.getenv('EXCHANGE_PASSWORD')
    server = os.getenv('EXCHANGE_SERVER')
    version = os.getenv('EXCHANGE_VERSION')
    output_dir = os.getenv('OUTPUT_DIR')
    timezone_name = os.getenv('TIMEZONE')
    days_ago = int(os.getenv('DAYS_AGO', 7))
    
    if not all([email, domain_username, password, server, version]):
        raise ValueError("Missing environment variables.")
    
    credentials = Credentials(username=domain_username, password=password)
    config = Configuration(server=server, credentials=credentials)
    account = Account(
        primary_smtp_address=email,
        credentials=credentials,
        autodiscover=True,
        access_type=DELEGATE,
        config=config
    )
    
    local_tz = pytz.timezone(timezone_name)
    time_frame = local_tz.localize(datetime.now() - timedelta(days=days_ago))
    
    return account, output_dir, time_frame


@app.route('/')
def index():
    try:
        # Get all HTML files in the output directory
        recent_emails = []
        for root, dirs, files in os.walk(EMAIL_DIR):
            for file in files:
                if file.endswith('.html'):
                    file_path = os.path.join(root, file)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()
                            # Extract metadata using regex
                            subject_match = re.search(r'<h1>Subject: (.*?)</h1>', content)
                            sender_match = re.search(r'<p><strong>Sender:</strong> (.*?)</p>', content)
                            datetime_match = re.search(r'<p><strong>Received:</strong> (.*?)</p>', content)
                            
                            # Get metadata or default values
                            subject = subject_match.group(1) if subject_match else 'No Subject'
                            sender = sender_match.group(1) if sender_match else 'Unknown Sender'
                            datetime_str = datetime_match.group(1) if datetime_match else None
                            
                            if datetime_str:
                                try:
                                    dt = datetime.fromisoformat(datetime_str.replace('Z', '+00:00'))
                                except ValueError:
                                    dt = datetime.fromtimestamp(os.path.getmtime(file_path))
                            else:
                                dt = datetime.fromtimestamp(os.path.getmtime(file_path))
                            
                            recent_emails.append({
                                'subject': subject,
                                'sender': sender,
                                'datetime_received': dt,
                                'path': os.path.relpath(file_path, EMAIL_DIR)
                            })
                    except Exception as e:
                        logging.error(f"Error reading email file {file_path}: {str(e)}")
        
        # Sort emails by datetime_received in descending order
        recent_emails.sort(key=lambda x: x['datetime_received'], reverse=True)
        
        # Take only the 10 most recent emails
        recent_emails = recent_emails[:15]
        
        # Format the datetime for display
        for email in recent_emails:
            email['datetime_received'] = email['datetime_received'].strftime('%m/%d/%Y %I:%M %p')
        
        return render_template('index.html', emails=recent_emails)
    except Exception as e:
        logging.error(f"Error fetching recent emails: {str(e)}")
        return render_template('index.html', emails=[])

@app.route('/search')
def search():
    print()
    query = request.args.get('query', '').strip()
    if not query:
        return jsonify({"error": "No search query provided."}), 400
    
    print(f"Searching for: {query}")  # Debug print
    
    results = []
    try:
        for root, dirs, files in os.walk(EMAIL_DIR):
            for file in files:
                if file.endswith('.html'):
                    file_path = os.path.join(root, file)
                    # print(f"Checking file: {file_path}")  # Debug print
                    
                    relative_path = os.path.relpath(file_path, EMAIL_DIR)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()
                            if query.lower() in content.lower():
                                print(f"Found match in: {file_path}")  # Debug print
                                snippet_start = content.lower().find(query.lower())
                                snippet = content[max(0, snippet_start-50):snippet_start+150] + '...'
                                snippet = re.sub('<[^<]+?>', '', snippet)
                                results.append({
                                    'path': relative_path.replace('\\', '/'),
                                    'name': file,
                                    'snippet': snippet
                                })
                    except Exception as e:
                        print(f"Error reading file {file_path}: {str(e)}")  # Debug print
                        
        print(f"Found {len(results)} results")  # Debug print
        return jsonify({"results": results})
    
    except Exception as e:
        print(f"Search error: {str(e)}")  # Debug print
        return jsonify({"error": str(e)}), 500

@app.route('/view/<path:filename>')
def view(filename):
    try:
        # Ensure the path is within EMAIL_DIR
        full_path = os.path.join(EMAIL_DIR, filename)
        full_path = os.path.normpath(full_path)
        
        if not full_path.startswith(os.path.normpath(EMAIL_DIR)):
            abort(403)  # Forbidden if trying to access outside EMAIL_DIR
            
        if not os.path.exists(full_path):
            abort(404)  # Not found
            
        # Read and return the file content
        with open(full_path, 'r', encoding='utf-8') as f:
            content = f.read()
            return content
            
    except Exception as e:
        logging.error(f"Error reading email file: {str(e)}")
        abort(500)

@app.route('/attachments/<path:filename>')
def download_attachment(filename):
    """Route to download attachments."""
    ic()
    if '..' in filename or filename.startswith('/'):
        abort(400, description="Invalid file path.")
    
    # Split the path to get directory and file
    directory, file = os.path.split(filename)
    try:
        return send_from_directory(os.path.join(EMAIL_DIR, directory), file, as_attachment=True)
    except FileNotFoundError:
        abort(404, description="File not found.")

@app.route('/check-emails', methods=['POST'])
def check_emails():
    try:
        account, output_dir, time_frame = setup_exchange_connection()
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        # Process emails
        logging.info("Processing sent emails...")
        for _ in process_email(account, account.sent, output_dir, time_frame):
            pass
            
        logging.info("Processing inbox emails...")
        for _ in process_email(account, account.inbox, output_dir, time_frame):
            pass
            
        return jsonify({"success": True, "message": "Emails checked successfully"})
    except Exception as e:
        logging.error(f"Failed to check emails: {str(e)}")
        return jsonify({"success": False, "message": str(e)}), 500

if __name__ == '__main__':
    # When running the Flask app, you might also want to process emails
    # Uncomment the following lines if you want to process emails on startup
    
    try:
        account, output_dir, time_frame = setup_exchange_connection()
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        logging.info("Processing sent emails...")
        for _ in process_email(account, account.sent, output_dir, time_frame):
            pass
        logging.info("Processing inbox emails...")
        for _ in process_email(account, account.inbox, output_dir, time_frame):
            pass
    except Exception as e:
        logging.error(f"Failed to setup Exchange connection or process emails: {str(e)}")
    
    print()
    app.run(debug=True, host='0.0.0.0')