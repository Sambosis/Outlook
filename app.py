# app.py    
import re
import os
import logging
import json
from json.decoder import JSONDecodeError
from flask import Flask, render_template, request, jsonify, send_from_directory, abort
from datetime import datetime, timedelta
from exchangelib import Credentials, Account, DELEGATE, Message, FileAttachment, ItemAttachment, Configuration
from exchangelib.errors import ErrorTooManyObjectsOpened
import pytz
from pathlib import Path
from icecream import ic 
from waitress import serve
from dotenv import load_dotenv
import sys

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Get environment variables with validation
def get_env_var(var_name):
    # Get the absolute path of the .env file
    env_path = Path(__file__).resolve().parent / '.env'
    # logger.debug(f"Loading .env file from: {env_path}")
    
    # Force reload of .env file
    load_dotenv(env_path, override=True)
    
    value = os.getenv(var_name)
    logger.debug(f"Variable {var_name} = {value}")
    
    if value is None:
        raise ValueError(f"Missing environment variable: {var_name}")
    return value.strip("'\"")  # Remove any quotes

try:
    logger.debug("Starting environment variable loading...")
    EXCHANGE_EMAIL = get_env_var('EXCHANGE_EMAIL')
    EXCHANGE_DOMAIN_USERNAME = get_env_var('EXCHANGE_DOMAIN_USERNAME')
    EXCHANGE_PASSWORD = get_env_var('EXCHANGE_PASSWORD')
    EXCHANGE_SERVER = get_env_var('EXCHANGE_SERVER')
    EXCHANGE_VERSION = get_env_var('EXCHANGE_VERSION')
    OUTPUT_DIR = get_env_var('OUTPUT_DIR')
    TIMEZONE = get_env_var('TIMEZONE')
    DAYS_AGO = int(get_env_var('DAYS_AGO'))
    logger.debug("Finished loading environment variables")
except ValueError as e:
    logger.error(f"Environment configuration error: {str(e)}")
    raise

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__, static_folder='static')

# Directory where emails are stored
EMAIL_DIR = 'gpg2/'  # Ensure this path is correct
# sys.exit(1)
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
# @pysnooper.snoop("ouput.log")
def process_email_item(account, item, output_dir):
    """Process a single email item."""
    recipient_name = item.to_recipients[0].name if item.to_recipients else 'Unknown_Recipient'
    subject = sanitize_filename(item.subject) if item.subject else 'No_Subject'
    date_str = item.datetime_received.strftime('%m-%d-%Y_%I-%M%p')

    # Define base name for files and attachments
    base_name = f"to_{recipient_name} - {subject} - {date_str}"

    # Construct paths
    email_filename = os.path.join(EMAIL_DIR, f"{base_name}.html")
    attachment_dir = os.path.join(EMAIL_DIR, f"{base_name}_attachments")

    try:
        with open(email_filename, 'w', encoding='utf-8') as f:
            f.write("<html><body>\n")
            f.write(f"<h1>Subject: {item.subject}</h1>\n")
            f.write(f"<p><strong>Received:</strong> {item.datetime_received}</p>\n")
            f.write(f"<p><strong>Sender:</strong> {item.sender.email_address}</p>\n")
            to_addresses = ', '.join([r.email_address for r in item.to_recipients if r.email_address])
            f.write(f"<p><strong>To:</strong> {to_addresses}</p>\n")
            f.write("<p><strong>Body:</strong></p>\n")
            f.write(f"{item.body}\n")
            f.write("</body></html>\n")
    except Exception as e:
        logging.error(f"Error writing email file {email_filename}: {str(e)}")

    # If attachments exist, store them under the base_name_attachments folder
    if item.attachments:
        Path(attachment_dir).mkdir(parents=True, exist_ok=True)
        for attachment in item.attachments:
            if isinstance(attachment, FileAttachment):
                safe_attachment_name = sanitize_filename(attachment.name)
                attachment_path = os.path.join(attachment_dir, safe_attachment_name)
                try:
                    with open(attachment_path, 'wb') as f:
                        f.write(attachment.content)
                except Exception as e:
                    logging.error(f"Error saving attachment {attachment.name}: {str(e)}")
            elif isinstance(attachment, ItemAttachment):
                try:
                    attached_item = attachment.item
                    attached_subject = sanitize_filename(attached_item.subject)
                    attached_file = os.path.join(
                        attachment_dir,
                        f"attached_email_{attached_subject}_{attached_item.datetime_received.strftime('%Y%m%d%H%M%S')}.html"
                    )
                    with open(attached_file, 'w', encoding='utf-8') as f:
                        f.write("<html><body>\n")
                        f.write(f"<h1>Subject: {attached_item.subject}</h1>\n")
                        f.write(f"<p><strong>Received:</strong> {attached_item.datetime_received}</p>\n")
                        f.write(f"<p><strong>Sender:</strong> {attached_item.sender.email_address}</p>\n")
                        f.write("<p><strong>Body:</strong></p>\n")
                        f.write(f"{attached_item.body}\n")
                        f.write("</body></html>\n")
                except Exception as e:
                    logging.error(f"Error saving attached email: {str(e)}")

    return base_name

def setup_exchange_connection():
    """Setup Exchange connection using environment variables."""
    email = EXCHANGE_EMAIL
    domain_username = EXCHANGE_DOMAIN_USERNAME
    password = EXCHANGE_PASSWORD
    server = EXCHANGE_SERVER
    version = EXCHANGE_VERSION
    output_dir = OUTPUT_DIR
    timezone_name = TIMEZONE
    days_ago = DAYS_AGO
    
    if not all([email, domain_username, password, server, version]):
        logging.error("One or more environment variables are missing.")
        raise ValueError("Missing environment variables.")
    
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
        
        # Take only the 100 most recent emails
        recent_emails = recent_emails[:100]
        
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
                    relative_path = os.path.relpath(file_path, EMAIL_DIR).replace('\\', '/')
                    
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()
                            if query.lower() in content.lower():
                                # Extract metadata using regex
                                subject_match = re.search(r'<h1>Subject: (.*?)</h1>', content)
                                sender_match = re.search(r'<p><strong>Sender:</strong> (.*?)</p>', content)
                                datetime_match = re.search(r'<p><strong>Received:</strong> (.*?)</p>', content)
                                
                                # Get metadata or default values
                                subject = subject_match.group(1) if subject_match else 'No Subject'
                                sender = sender_match.group(1) if sender_match else 'Unknown Sender'
                                datetime_str = datetime_match.group(1) if datetime_match else None
                                
                                dt = None
                                if datetime_str:
                                    try:
                                        dt = datetime.fromisoformat(datetime_str.replace('Z', '+00:00'))
                                    except ValueError:
                                        dt = datetime.fromtimestamp(os.path.getmtime(file_path))
                                else:
                                    dt = datetime.fromtimestamp(os.path.getmtime(file_path))

                                # Extract snippet
                                snippet_start = content.lower().find(query.lower())
                                snippet = content[max(0, snippet_start-50):snippet_start+150] + '...'
                                snippet = re.sub('<[^<]+?>', '', snippet) # Basic HTML tag removal for snippet

                                results.append({
                                    'path': relative_path,
                                    'subject': subject,
                                    'sender': sender,
                                    'datetime_obj': dt, # Store datetime object for sorting
                                    'snippet': snippet # Keep snippet if needed, or remove if unused
                                })
                    except Exception as e:
                        print(f"Error reading file {file_path}: {str(e)}")

        # Sort results by datetime object in descending order
        results.sort(key=lambda x: x['datetime_obj'], reverse=True)
        
        # Format datetime for display and remove the object
        for result in results:
            result['datetime_received'] = result['datetime_obj'].strftime('%m/%d/%Y %I:%M %p')
            del result['datetime_obj']
            
        return jsonify({"results": results})
    
    except Exception as e:
        print(f"Search error: {str(e)}")  # Debug print
        return jsonify({"error": str(e)}), 500

@app.route('/view/<path:filename>')
def view(filename):
    try:
        full_path = os.path.join(EMAIL_DIR, filename)
        full_path = os.path.normpath(full_path)
        
        if not full_path.startswith(os.path.normpath(EMAIL_DIR)):
            abort(403)
            
        if not os.path.exists(full_path):
            abort(404)
            
        with open(full_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # Extract attachment directory name based on email filename
            attachment_dir_match = re.match(r'(.*)\.html$', filename)
            if attachment_dir_match:
                base_name = attachment_dir_match.group(1)
                attachment_dir = f"{base_name}_attachments/"
                # Optionally, verify if attachment directory exists
                attachments_path = os.path.join(EMAIL_DIR, attachment_dir)
                if os.path.exists(attachments_path):
                    # You can modify the content to include attachment links if needed
                    pass
            return content
            
    except Exception as e:
        logging.error(f"Error reading email file: {str(e)}")
        abort(500)

@app.route('/gpg2/<path:filename>')
def download_attachment(filename):
    """Enhanced route to download attachments with security checks"""
    if '..' in filename or filename.startswith('/'):
        abort(400, description="Invalid file path")
        
    full_path = os.path.join(EMAIL_DIR, filename)
    full_path = os.path.normpath(full_path)
    
    if not full_path.startswith(os.path.normpath(EMAIL_DIR)):
        abort(403)  # Forbidden if trying to access outside EMAIL_DIR
        
    try:
        directory = os.path.dirname(full_path)
        file = os.path.basename(full_path)
        return send_from_directory(directory, file, as_attachment=True)
    except FileNotFoundError:
        abort(404, description="Attachment not found")
    except Exception as e:
        logging.error(f"Error downloading attachment: {str(e)}")
        abort(500)

@app.route('/list-attachments/<path:email_path>')
def list_attachments(email_path):
    try:
        # Remove .html extension to get base path
        email_base = email_path.rsplit('.', 1)[0]
        
        attachment_dir = f"{email_base}_attachments/"
        attachment_full_path = os.path.join(EMAIL_DIR, attachment_dir)
        
        if os.path.exists(attachment_full_path):
            attachments = []
            for file in os.listdir(attachment_full_path):
                file_path = os.path.join(attachment_full_path, file)
                if os.path.isfile(file_path):
                    attachments.append({
                        'filename': file,
                        'path': f'/gpg2/{attachment_dir}{file}',
                        'size': os.path.getsize(file_path)
                    })
            logging.debug(f"Found {len(attachments)} attachments")
            return jsonify({'attachments': attachments})
        
        return jsonify({'attachments': []})
    except Exception as e:
        logging.error(f"Error listing attachments: {str(e)}")
        return jsonify({'error': str(e), 'attachments': []}), 500

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
    except JSONDecodeError as e:
        logging.error(f"JSON decode error: {str(e)}")
        return jsonify({"success": False, "message": f"JSON decode error: {str(e)}"}), 500
    except Exception as e:
        logging.error(f"Failed to check emails: {str(e)}")
        return jsonify({"success": False, "message": str(e)}), 500

@app.route('/gpg2/<path:filename>')
def serve_attachment(filename):
    """Serve attachment files securely."""
    try:
        full_path = os.path.join(EMAIL_DIR, filename)
        full_path = os.path.normpath(full_path)
        
        if not full_path.startswith(os.path.normpath(EMAIL_DIR)):
            abort(403)  # Forbidden if trying to access outside EMAIL_DIR
            
        if not os.path.exists(full_path):
            abort(404)
        
        directory = os.path.dirname(full_path)
        file = os.path.basename(full_path)
        return send_from_directory(directory, file, as_attachment=True)
    except Exception as e:
        logging.error(f"Error serving attachment {filename}: {str(e)}")
        abort(500)

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
    # app.run(debug=True)
    serve(app, host='127.0.0.1', port=8080)
