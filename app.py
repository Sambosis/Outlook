# app.py
import re
import os
import logging
import json
from json.decoder import JSONDecodeError
from flask import Flask, render_template, request, jsonify, abort, send_file, Response
from datetime import datetime, timedelta
from exchangelib import Credentials, Account, DELEGATE, Message, FileAttachment, ItemAttachment, Configuration
from exchangelib.errors import ErrorTooManyObjectsOpened
import pytz
from pathlib import Path
from waitress import serve
from dotenv import load_dotenv
import sys
from sqlalchemy import text
from api_wrapper import ensure_json_response, json_response
from database import get_db_connection, create_tables

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Get environment variables with validation
def get_env_var(var_name):
    env_path = Path(__file__).resolve().parent / '.env'
    load_dotenv(env_path, override=True)
    value = os.getenv(var_name)
    if value is None:
        raise ValueError(f"Missing environment variable: {var_name}")
    return value.strip("'\"")

try:
    logger.debug("Starting environment variable loading...")
    EXCHANGE_EMAIL = get_env_var('EXCHANGE_EMAIL')
    EXCHANGE_DOMAIN_USERNAME = get_env_var('EXCHANGE_DOMAIN_USERNAME')
    EXCHANGE_PASSWORD = get_env_var('EXCHANGE_PASSWORD')
    EXCHANGE_SERVER = get_env_var('EXCHANGE_SERVER')
    EXCHANGE_VERSION = get_env_var('EXCHANGE_VERSION')
    TIMEZONE = get_env_var('TIMEZONE')
    DAYS_AGO = int(get_env_var('DAYS_AGO'))
    logger.debug("Finished loading environment variables")
except ValueError as e:
    logger.error(f"Environment configuration error: {str(e)}")
    raise

app = Flask(__name__, static_folder='static')

@app.errorhandler(Exception)
def handle_exception(e):
    code = getattr(e, 'code', 500)
    return jsonify({
        "success": False,
        "message": str(e),
        "error": type(e).__name__
    }), code

def sanitize_filename(filename):
    invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
    filename = re.sub(invalid_chars, '_', filename)
    return filename.rstrip('. ')

def process_email(account, email_folder, time_frame):
    try:
        for item in email_folder.filter(datetime_received__gte=time_frame).order_by('-datetime_received'):
            if isinstance(item, Message):
                yield process_email_item(item)
    except ErrorTooManyObjectsOpened as e:
        logging.error(f"Too many objects error: {e}")

def process_email_item(item):
    conn = None
    try:
        conn = get_db_connection()
        trans = conn.begin()

        # Insert email
        insert_email_query = text("""
            INSERT INTO emails (subject, sender, recipients, body, received_at)
            VALUES (:subject, :sender, :recipients, :body, :received_at)
            RETURNING id;
        """)

        to_addresses = ', '.join([r.email_address for r in item.to_recipients if r.email_address])

        result = conn.execute(insert_email_query, {
            'subject': item.subject,
            'sender': item.sender.email_address,
            'recipients': to_addresses,
            'body': item.body,
            'received_at': item.datetime_received
        })
        email_id = result.fetchone()[0]

        # Handle attachments
        for attachment in item.attachments:
            if isinstance(attachment, FileAttachment):
                insert_attachment_query = text("""
                    INSERT INTO attachments (email_id, filename, content)
                    VALUES (:email_id, :filename, :content);
                """)
                conn.execute(insert_attachment_query, {
                    'email_id': email_id,
                    'filename': sanitize_filename(attachment.name),
                    'content': attachment.content
                })

        trans.commit()
        return email_id
    except Exception as e:
        if 'trans' in locals() and trans:
            trans.rollback()
        logging.error(f"Error processing email item: {e}")
    finally:
        if conn:
            conn.close()

def setup_exchange_connection():
    credentials = Credentials(username=EXCHANGE_DOMAIN_USERNAME, password=EXCHANGE_PASSWORD)
    config = Configuration(server=EXCHANGE_SERVER, credentials=credentials)
    account = Account(
        primary_smtp_address=EXCHANGE_EMAIL,
        credentials=credentials,
        autodiscover=True,
        access_type=DELEGATE,
        config=config
    )
    local_tz = pytz.timezone(TIMEZONE)
    time_frame = local_tz.localize(datetime.now() - timedelta(days=DAYS_AGO))
    return account, time_frame

@app.route('/')
def index():
    conn = None
    try:
        conn = get_db_connection()
        query = text("SELECT id, subject, sender, received_at FROM emails ORDER BY received_at DESC LIMIT 100")
        result = conn.execute(query)
        emails = [{
            'id': row[0],
            'subject': row[1],
            'sender': row[2],
            'datetime_received': row[3].strftime('%m/%d/%Y %I:%M %p')
        } for row in result]
        return render_template('index.html', emails=emails)
    except Exception as e:
        logging.error(f"Error fetching recent emails: {e}")
        return render_template('index.html', emails=[])
    finally:
        if conn:
            conn.close()

@app.route('/search')
@ensure_json_response
def search():
    query = request.args.get('query', '').strip()
    if not query:
        return jsonify({"error": "No search query provided."}), 400
    
    conn = None
    try:
        conn = get_db_connection()
        search_query = text("""
            SELECT id, subject, sender, received_at, body
            FROM emails
            WHERE subject ILIKE :query OR body ILIKE :query OR sender ILIKE :query
            ORDER BY received_at DESC
        """)
        results = conn.execute(search_query, {'query': f'%{query}%'}).fetchall()
        
        search_results = []
        for row in results:
            body_content = row[4]
            snippet_start = body_content.lower().find(query.lower())
            snippet = body_content[max(0, snippet_start-50):snippet_start+150] + '...'
            snippet = re.sub('<[^<]+?>', '', snippet)
            
            search_results.append({
                'id': row[0],
                'subject': row[1],
                'sender': row[2],
                'datetime_received': row[3].strftime('%m/%d/%Y %I:%M %p'),
                'snippet': snippet
            })
        return jsonify({"results": search_results})
    except Exception as e:
        logging.error(f"Search error: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        if conn:
            conn.close()

@app.route('/view/<int:email_id>')
def view(email_id):
    conn = None
    try:
        conn = get_db_connection()
        query = text("SELECT body FROM emails WHERE id = :id")
        result = conn.execute(query, {'id': email_id}).fetchone()
        if result:
            return result[0]
        abort(404)
    except Exception as e:
        logging.error(f"Error reading email: {e}")
        abort(500)
    finally:
        if conn:
            conn.close()

@app.route('/download_attachment/<int:attachment_id>')
def download_attachment(attachment_id):
    conn = None
    try:
        conn = get_db_connection()
        query = text("SELECT filename, content FROM attachments WHERE id = :id")
        result = conn.execute(query, {'id': attachment_id}).fetchone()
        if result:
            filename, content = result
            return Response(
                content,
                mimetype='application/octet-stream',
                headers={'Content-Disposition': f'attachment;filename={filename}'}
            )
        abort(404)
    except Exception as e:
        logging.error(f"Error downloading attachment: {e}")
        abort(500)
    finally:
        if conn:
            conn.close()

@app.route('/list-attachments/<int:email_id>')
@ensure_json_response
def list_attachments(email_id):
    conn = None
    try:
        conn = get_db_connection()
        query = text("SELECT id, filename, length(content) as size FROM attachments WHERE email_id = :email_id")
        result = conn.execute(query, {'email_id': email_id})
        attachments = [{
            'id': row[0],
            'filename': row[1],
            'size': row[2]
        } for row in result]
        return {'attachments': attachments}
    except Exception as e:
        logging.error(f"Error listing attachments: {e}")
        return json_response(success=False, message=f"Error listing attachments: {e}", status_code=500)
    finally:
        if conn:
            conn.close()

@app.route('/check-emails', methods=['POST'])
@ensure_json_response
def check_emails():
    try:
        account, time_frame = setup_exchange_connection()
        
        processed_sent = 0
        processed_inbox = 0
        
        logging.info("Processing sent emails...")
        for _ in process_email(account, account.sent, time_frame):
            processed_sent += 1
            
        logging.info("Processing inbox emails...")
        for _ in process_email(account, account.inbox, time_frame):
            processed_inbox += 1
            
        return json_response(
            success=True, 
            message=f"Emails checked successfully. Processed {processed_sent} sent and {processed_inbox} inbox emails.",
            data={"sent": processed_sent, "inbox": processed_inbox}
        )
    except Exception as e:
        logging.error(f"Failed to check emails: {e}")
        return json_response(success=False, message=f"Failed to check emails: {e}", status_code=500)

if __name__ == '__main__':
    create_tables()
    try:
        account, time_frame = setup_exchange_connection()
        logging.info("Processing sent emails on startup...")
        for _ in process_email(account, account.sent, time_frame):
            pass
        logging.info("Processing inbox emails on startup...")
        for _ in process_email(account, account.inbox, time_frame):
            pass
    except Exception as e:
        logging.error(f"Failed to process emails on startup: {e}")
    
    serve(app, host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
