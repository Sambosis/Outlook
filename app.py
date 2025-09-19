# app.py
import logging
import os
import re
import zipfile
from datetime import datetime, timedelta
from io import BytesIO

import pytz
from exchangelib import (
    Account,
    Configuration,
    Credentials,
    DELEGATE,
    FileAttachment,
    ItemAttachment,
    Message,
)
from exchangelib.errors import ErrorTooManyObjectsOpened
from flask import (
    Flask,
    abort,
    jsonify,
    render_template,
    request,
    send_file,
    url_for,
)
from dotenv import load_dotenv
from sqlalchemy import or_
from sqlalchemy.orm import selectinload
from waitress import serve

from api_wrapper import ensure_json_response, json_response
from db import SessionLocal, init_db
from models import Attachment, Email

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


# Get environment variables with validation
def get_env_var(var_name):
    env_path = os.path.join(os.path.dirname(__file__), '.env')
    load_dotenv(env_path, override=True)

    value = os.getenv(var_name)
    logger.debug(f"Variable {var_name} = {value}")

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
    OUTPUT_DIR = os.getenv('OUTPUT_DIR')
    logger.debug("Finished loading environment variables")
except ValueError as e:
    logger.error(f"Environment configuration error: {str(e)}")
    raise

if OUTPUT_DIR:
    os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__, static_folder='static')
init_db()

LOCAL_TZ = pytz.timezone(TIMEZONE)
CID_PATTERN = re.compile(r'src=["\']cid:(.*?)["\']', re.IGNORECASE)


@app.errorhandler(Exception)
def handle_exception(e):
    code = getattr(e, 'code', 500)
    return jsonify({
        "success": False,
        "message": str(e),
        "error": type(e).__name__
    }), code


def sanitize_filename(filename):
    if not filename:
        return "unnamed"

    invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
    filename = re.sub(invalid_chars, '_', filename)
    filename = filename.rstrip('. ')
    if not filename:
        filename = 'unnamed'
    if len(filename) > 245:
        filename = filename[:245]
    return filename


def strip_html_tags(value: str) -> str:
    if not value:
        return ""
    return re.sub(r'<[^>]+>', '', value)


def ensure_utc(dt):
    if not dt:
        return None
    if dt.tzinfo is None:
        return pytz.utc.localize(dt)
    return dt.astimezone(pytz.utc)


def format_datetime_for_display(dt):
    if not dt:
        return "Unknown"
    if dt.tzinfo is None:
        dt = pytz.utc.localize(dt)
    return dt.astimezone(LOCAL_TZ).strftime('%m/%d/%Y %I:%M %p')


def build_email_base_name(email: Email) -> str:
    recipient_name = sanitize_filename(email.primary_recipient or 'Unknown_Recipient')
    subject = sanitize_filename(email.subject or 'No_Subject')
    dt = email.datetime_received
    if dt:
        if dt.tzinfo is None:
            dt = pytz.utc.localize(dt)
        dt_str = dt.astimezone(LOCAL_TZ).strftime('%m-%d-%Y_%I-%M%p')
    else:
        dt_str = datetime.utcnow().strftime('%m-%d-%Y_%I-%M%p')
    return f"to_{recipient_name} - {subject} - {dt_str}"


def replace_cid_urls(body: str, attachments):
    if not body:
        return ""

    def cid_replacer(match):
        cid = match.group(1).strip()
        cid_clean = cid.strip('<>')
        for attachment in attachments:
            if attachment.content_id:
                attachment_cid = attachment.content_id.strip('<>')
                if attachment_cid == cid_clean:
                    return f'src="{url_for("download_attachment", attachment_id=attachment.id)}"'
        return match.group(0)

    return CID_PATTERN.sub(cid_replacer, body)


def generate_email_html(email: Email) -> str:
    body = replace_cid_urls(email.body or '', email.attachments)
    subject = email.subject or 'No Subject'
    sender = email.sender or 'Unknown Sender'
    recipients = email.recipients or 'Unknown Recipients'
    received = format_datetime_for_display(email.datetime_received)

    html_parts = [
        '<html><body>',
        f'<h1>Subject: {subject}</h1>',
        f'<p><strong>Received:</strong> {received}</p>',
        f'<p><strong>Sender:</strong> {sender}</p>',
        f'<p><strong>To:</strong> {recipients}</p>',
        '<p><strong>Body:</strong></p>',
        body,
        '</body></html>'
    ]
    return '\n'.join(html_parts)


def get_message_identifier(item: Message):
    message_id = getattr(item, 'message_id', None)
    if message_id:
        return message_id
    item_id = getattr(item, 'item_id', None)
    if item_id:
        identifier = getattr(item_id, 'id', None)
        if identifier:
            return identifier
        return str(item_id)
    return None


def render_item_attachment(attached_item):
    subject = getattr(attached_item, 'subject', 'Attached Email') or 'Attached Email'
    subject_safe = sanitize_filename(subject)
    received = getattr(attached_item, 'datetime_received', None)
    if received:
        received_str = ensure_utc(received).strftime('%Y%m%d%H%M%S')
    else:
        received_str = datetime.utcnow().strftime('%Y%m%d%H%M%S')

    sender = getattr(getattr(attached_item, 'sender', None), 'email_address', 'Unknown Sender')
    recipients = []
    for recipient in getattr(attached_item, 'to_recipients', []) or []:
        address = getattr(recipient, 'email_address', None)
        name = getattr(recipient, 'name', None)
        if name and address:
            recipients.append(f"{name} <{address}>")
        elif address:
            recipients.append(address)
        elif name:
            recipients.append(name)

    html_parts = [
        '<html><body>',
        f'<h1>Subject: {subject}</h1>',
        f'<p><strong>Received:</strong> {received}</p>',
        f'<p><strong>Sender:</strong> {sender}</p>',
        f"<p><strong>To:</strong> {', '.join(recipients) if recipients else 'Unknown Recipients'}</p>",
        '<p><strong>Body:</strong></p>',
        str(getattr(attached_item, 'body', '') or ''),
        '</body></html>'
    ]

    filename = f"attached_email_{subject_safe}_{received_str}.html"
    return filename, '\n'.join(html_parts).encode('utf-8')


def process_email_item(session, item: Message, folder_name: str) -> bool:
    message_identifier = get_message_identifier(item)
    if message_identifier:
        existing = session.query(Email).filter(Email.message_id == message_identifier).first()
        if existing:
            return False

    recipients = []
    primary_recipient = None
    to_recipients = getattr(item, 'to_recipients', None) or []
    for recipient in to_recipients:
        address = getattr(recipient, 'email_address', None)
        name = getattr(recipient, 'name', None)
        if name and address:
            recipients.append(f"{name} <{address}>")
        elif address:
            recipients.append(address)
        elif name:
            recipients.append(name)
    if to_recipients:
        first_recipient = to_recipients[0]
        primary_recipient = getattr(first_recipient, 'name', None) or getattr(first_recipient, 'email_address', None)

    dt_received = ensure_utc(getattr(item, 'datetime_received', None))
    dt_store = dt_received.replace(tzinfo=None) if dt_received else datetime.utcnow()

    body_html = str(getattr(item, 'body', '') or '')
    body_plain = strip_html_tags(body_html)

    email = Email(
        message_id=message_identifier,
        subject=getattr(item, 'subject', None),
        sender=getattr(getattr(item, 'sender', None), 'email_address', None),
        recipients=', '.join(recipients) if recipients else None,
        primary_recipient=primary_recipient,
        folder=folder_name,
        datetime_received=dt_store,
        body=body_html,
        body_plain=body_plain,
    )

    session.add(email)
    session.flush()

    for attachment in getattr(item, 'attachments', []) or []:
        if isinstance(attachment, FileAttachment):
            data = attachment.content or b''
            filename = sanitize_filename(getattr(attachment, 'name', None) or 'attachment')
            new_attachment = Attachment(
                email_id=email.id,
                filename=filename,
                content_type=getattr(attachment, 'content_type', None),
                size=len(data) if data else 0,
                content_id=getattr(attachment, 'content_id', None),
                data=data,
            )
            session.add(new_attachment)
        elif isinstance(attachment, ItemAttachment):
            try:
                attached_item = attachment.item
                filename, data = render_item_attachment(attached_item)
                session.add(
                    Attachment(
                        email_id=email.id,
                        filename=sanitize_filename(filename),
                        content_type='text/html',
                        size=len(data),
                        data=data,
                    )
                )
            except Exception as exc:
                logger.error(f"Error saving attached email: {exc}")

    return True


def process_email_folder(session, email_folder, time_frame, folder_name):
    processed = 0
    try:
        items = email_folder.filter(datetime_received__gte=time_frame).order_by('-datetime_received')
        for item in items:
            if not isinstance(item, Message):
                continue
            try:
                created = process_email_item(session, item, folder_name)
                if created:
                    session.commit()
                    processed += 1
            except Exception as exc:
                session.rollback()
                logger.error(f"Error processing email {getattr(item, 'subject', 'unknown')}: {exc}")
    except ErrorTooManyObjectsOpened as exc:
        logger.error(f"Too many objects error: {exc}")
    except Exception as exc:
        logger.error(f"Unexpected error while processing folder: {exc}")
    return processed


def setup_exchange_connection():
    email = EXCHANGE_EMAIL
    domain_username = EXCHANGE_DOMAIN_USERNAME
    password = EXCHANGE_PASSWORD
    server = EXCHANGE_SERVER
    version = EXCHANGE_VERSION
    timezone_name = TIMEZONE
    days_ago = DAYS_AGO

    if not all([email, domain_username, password, server, version]):
        logger.error("One or more environment variables are missing.")
        raise ValueError("Missing environment variables.")

    credentials = Credentials(username=domain_username, password=password)
    config = Configuration(server=server, credentials=credentials)
    account = Account(
        primary_smtp_address=email,
        credentials=credentials,
        autodiscover=True,
        access_type=DELEGATE,
        config=config,
    )

    local_tz = pytz.timezone(timezone_name)
    time_frame = local_tz.localize(datetime.now() - timedelta(days=days_ago))

    return account, time_frame


@app.route('/')
def index():
    try:
        with SessionLocal() as session:
            emails = (
                session.query(Email)
                .order_by(Email.datetime_received.desc())
                .limit(100)
                .all()
            )
            recent_emails = [
                {
                    'id': email.id,
                    'subject': email.subject or 'No Subject',
                    'sender': email.sender or 'Unknown Sender',
                    'datetime_received': format_datetime_for_display(email.datetime_received),
                }
                for email in emails
            ]
        return render_template('index.html', emails=recent_emails)
    except Exception as exc:
        logger.error(f"Error fetching recent emails: {exc}")
        return render_template('index.html', emails=[])


@app.route('/search')
@ensure_json_response
def search():
    query = request.args.get('query', '').strip()
    if not query:
        return {"error": "No search query provided."}, 400

    with SessionLocal() as session:
        results = (
            session.query(Email)
            .filter(
                or_(
                    Email.subject.ilike(f'%{query}%'),
                    Email.sender.ilike(f'%{query}%'),
                    Email.recipients.ilike(f'%{query}%'),
                    Email.body_plain.ilike(f'%{query}%'),
                )
            )
            .order_by(Email.datetime_received.desc())
            .limit(100)
            .all()
        )

    response = []
    for email in results:
        body_text = email.body_plain or ''
        index = body_text.lower().find(query.lower()) if query else -1
        snippet = ''
        if index != -1:
            start = max(0, index - 50)
            end = min(len(body_text), index + 150)
            snippet = body_text[start:end] + ('...' if end < len(body_text) else '')

        response.append({
            'id': email.id,
            'subject': email.subject or 'No Subject',
            'sender': email.sender or 'Unknown Sender',
            'datetime_received': format_datetime_for_display(email.datetime_received),
            'snippet': snippet,
        })

    return {"results": response}


@app.route('/email/<int:email_id>')
def view_email(email_id):
    with SessionLocal() as session:
        email = (
            session.query(Email)
            .options(selectinload(Email.attachments))
            .filter(Email.id == email_id)
            .first()
        )
        if not email:
            abort(404)
        return generate_email_html(email)


@app.route('/attachments/<int:email_id>')
@ensure_json_response
def list_attachments(email_id):
    with SessionLocal() as session:
        email = (
            session.query(Email)
            .options(selectinload(Email.attachments))
            .filter(Email.id == email_id)
            .first()
        )
        if not email:
            return {"success": False, "message": "Email not found"}, 404

        attachments = [
            {
                'id': attachment.id,
                'filename': attachment.filename,
                'size': attachment.size or 0,
                'download_url': url_for('download_attachment', attachment_id=attachment.id),
            }
            for attachment in email.attachments
        ]

    return {"attachments": attachments}


@app.route('/attachments/<int:attachment_id>/download')
def download_attachment(attachment_id):
    with SessionLocal() as session:
        attachment = session.get(Attachment, attachment_id)
        if not attachment:
            abort(404)

        data = attachment.data or b''
        file_stream = BytesIO(data)
        file_stream.seek(0)

        return send_file(
            file_stream,
            as_attachment=True,
            download_name=attachment.filename or 'attachment',
            mimetype=attachment.content_type or 'application/octet-stream',
        )


@app.route('/check-emails', methods=['POST'])
@ensure_json_response
def check_emails():
    try:
        account, time_frame = setup_exchange_connection()
    except Exception as exc:
        logger.error(f"Failed to setup Exchange connection: {exc}")
        return json_response(success=False, message=f"Failed to setup Exchange connection: {str(exc)}", status_code=500)

    with SessionLocal() as session:
        processed_sent = process_email_folder(session, account.sent, time_frame, 'sent')
        processed_inbox = process_email_folder(session, account.inbox, time_frame, 'inbox')

    return json_response(
        success=True,
        message=f"Emails checked successfully. Processed {processed_sent} sent and {processed_inbox} inbox emails.",
        data={"sent": processed_sent, "inbox": processed_inbox},
    )


@app.route('/download-all-emails')
def download_all_emails():
    try:
        with SessionLocal() as session:
            emails = (
                session.query(Email)
                .options(selectinload(Email.attachments))
                .order_by(Email.datetime_received.desc())
                .all()
            )

            memory_file = BytesIO()
            with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for email in emails:
                    base_name = build_email_base_name(email)
                    zipf.writestr(f"{base_name}.html", generate_email_html(email))
                    if email.attachments:
                        for attachment in email.attachments:
                            filename = sanitize_filename(attachment.filename or 'attachment')
                            folder_name = f"{base_name}_attachments"
                            zipf.writestr(f"{folder_name}/{filename}", attachment.data or b'')

            memory_file.seek(0)

        timestamp = datetime.utcnow().strftime('%Y%m%d_%H%M%S')
        zip_filename = f"all_emails_{timestamp}.zip"
        return send_file(
            memory_file,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip',
        )
    except Exception as exc:
        logger.error(f"Error creating zip file: {exc}")
        abort(500)


if __name__ == '__main__':
    try:
        account, time_frame = setup_exchange_connection()
        with SessionLocal() as session:
            process_email_folder(session, account.sent, time_frame, 'sent')
            process_email_folder(session, account.inbox, time_frame, 'inbox')
    except Exception as exc:
        logger.error(f"Failed to setup Exchange connection or process emails: {exc}")

    serve(app, host='127.0.0.1', port=8080)
