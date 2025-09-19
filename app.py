# app.py    
import io
import logging
import os
import re
import zipfile
from contextlib import contextmanager
from datetime import datetime, timedelta
from html import escape
from json.decoder import JSONDecodeError
from pathlib import Path

import pytz
from dotenv import load_dotenv
from exchangelib import Account, Configuration, Credentials, DELEGATE, FileAttachment, ItemAttachment, Message
from exchangelib.errors import ErrorTooManyObjectsOpened
from flask import (
    Flask,
    abort,
    jsonify,
    render_template,
    request,
    send_file,
)
from werkzeug.exceptions import HTTPException
from sqlalchemy import (
    Column,
    DateTime,
    ForeignKey,
    Integer,
    LargeBinary,
    String,
    Text,
    create_engine,
    inspect,
    or_,
    text,
)
from sqlalchemy.orm import (
    declarative_base,
    relationship,
    scoped_session,
    selectinload,
    sessionmaker,
)
from waitress import serve

from api_wrapper import ensure_json_response, json_response

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

# Database setup
DATABASE_URL = os.getenv('DATABASE_URL', 'sqlite:///emails.db')
if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)

engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = scoped_session(sessionmaker(bind=engine, autoflush=False, autocommit=False, expire_on_commit=False))
Base = declarative_base()


def utcnow():
    return datetime.now(pytz.UTC)


class Email(Base):
    __tablename__ = 'emails'

    id = Column(Integer, primary_key=True)
    message_id = Column(String(255), unique=True, index=True)
    subject = Column(Text)
    sender = Column(String(255))
    recipients = Column(Text)
    datetime_received = Column(DateTime(timezone=True))
    body = Column(Text)
    created_at = Column(DateTime(timezone=True), default=utcnow)

    attachments = relationship('Attachment', back_populates='email', cascade='all, delete-orphan')


class Attachment(Base):
    __tablename__ = 'attachments'

    id = Column(Integer, primary_key=True)
    email_id = Column(Integer, ForeignKey('emails.id', ondelete='CASCADE'), index=True)
    filename = Column(Text)
    content_type = Column(String(255))
    data = Column(LargeBinary)
    size = Column(Integer)
    created_at = Column(DateTime(timezone=True), default=utcnow)

    email = relationship('Email', back_populates='attachments')


Base.metadata.create_all(engine)


def ensure_database_schema(target_engine):
    """Ensure required columns exist on legacy databases without migrations."""

    inspector = inspect(target_engine)

    def ensure_created_at(table_name):
        if table_name not in inspector.get_table_names():
            return

        columns = {column['name'] for column in inspector.get_columns(table_name)}
        if 'created_at' in columns:
            return

        with target_engine.begin() as connection:
            dialect_name = target_engine.dialect.name
            if dialect_name == 'postgresql':
                connection.execute(
                    text(
                        f"ALTER TABLE {table_name} "
                        "ADD COLUMN created_at TIMESTAMPTZ DEFAULT timezone('UTC', now())"
                    )
                )
                connection.execute(
                    text(
                        f"UPDATE {table_name} "
                        "SET created_at = timezone('UTC', now()) "
                        "WHERE created_at IS NULL"
                    )
                )
            elif dialect_name == 'sqlite':
                connection.execute(
                    text(
                        f"ALTER TABLE {table_name} "
                        "ADD COLUMN created_at TIMESTAMP"
                    )
                )
                connection.execute(
                    text(
                        f"UPDATE {table_name} "
                        "SET created_at = CURRENT_TIMESTAMP "
                        "WHERE created_at IS NULL"
                    )
                )
            else:
                connection.execute(
                    text(
                        f"ALTER TABLE {table_name} "
                        "ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP"
                    )
                )
                connection.execute(
                    text(
                        f"UPDATE {table_name} "
                        "SET created_at = CURRENT_TIMESTAMP "
                        "WHERE created_at IS NULL"
                    )
                )

    ensure_created_at('emails')
    ensure_created_at('attachments')


ensure_database_schema(engine)


@contextmanager
def session_scope():
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


app = Flask(__name__, static_folder='static')

@app.errorhandler(Exception)
def handle_exception(e):
    code = getattr(e, 'code', 500)
    # Optionally log the full traceback here if needed
    return jsonify({
        "success": False,
        "message": str(e),
        "error": type(e).__name__
    }), code

def sanitize_filename(filename):
    """Sanitize the filename by removing or replacing invalid characters."""
    invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
    filename = re.sub(invalid_chars, '_', filename or '')
    filename = filename.rstrip('. ')
    if not filename:
        filename = 'unnamed'
    if len(filename) > 245:
        filename = filename[:245]
    return filename


def build_email_html(email_record):
    """Create an HTML representation of an email stored in the database."""
    subject = email_record.subject or 'No Subject'
    received = email_record.datetime_received.isoformat() if email_record.datetime_received else 'Unknown'
    sender = email_record.sender or 'Unknown Sender'
    recipients = email_record.recipients or 'Unknown'
    body = email_record.body or ''

    if body.strip():
        if re.search(r'<[^>]+>', body):
            body_html = body
        else:
            body_html = f"<pre>{escape(body)}</pre>"
    else:
        body_html = '<p><em>No body content</em></p>'

    return (
        "<html><body\n"
        f"<h1>Subject: {escape(subject)}</h1>\n"
        f"<p><strong>Received:</strong> {escape(received)}</p>\n"
        f"<p><strong>Sender:</strong> {escape(sender)}</p>\n"
        f"<p><strong>To:</strong> {escape(recipients)}</p>\n"
        "<p><strong>Body:</strong></p>\n"
        f"{body_html}\n"
        "</body></html>\n"
    )



def get_message_identifier(item):
    identifier = getattr(item, 'message_id', None)
    if identifier:
        return identifier
    return getattr(item, 'item_id', None)


def process_email(account, email_folder, session, time_frame):
    """Process emails in the specified folder and persist them to the database."""
    try:
        queryset = email_folder.filter(datetime_received__gte=time_frame).order_by('-datetime_received')
        for item in queryset:
            if isinstance(item, Message):
                yield process_email_item(item, session)
    except ErrorTooManyObjectsOpened as e:
        logging.error(f"Too many objects error: {e}")


def process_email_item(item, session):
    """Persist a single email item and its attachments."""
    message_identifier = get_message_identifier(item)
    if not message_identifier:
        message_identifier = f"{item.subject}-{item.datetime_received}-{getattr(item, 'sender', '')}"

    existing = session.query(Email).filter(Email.message_id == message_identifier).first()
    if existing:
        return False

    recipients = ', '.join([
        r.email_address or r.name
        for r in (item.to_recipients or [])
        if getattr(r, 'email_address', None) or getattr(r, 'name', None)
    ])

    email_record = Email(
        message_id=message_identifier,
        subject=item.subject,
        sender=item.sender.email_address if item.sender else None,
        recipients=recipients,
        datetime_received=item.datetime_received,
        body=str(item.body) if item.body is not None else None,
    )
    session.add(email_record)
    session.flush()

    if getattr(item, 'attachments', None):
        for attachment in item.attachments:
            try:
                if isinstance(attachment, FileAttachment):
                    data = attachment.content or b''
                    filename = sanitize_filename(attachment.name)
                    session.add(
                        Attachment(
                            email_id=email_record.id,
                            filename=filename,
                            content_type=getattr(attachment, 'content_type', None),
                            data=data,
                            size=len(data),
                        )
                    )
                elif isinstance(attachment, ItemAttachment):
                    attached_item = attachment.item
                    attached_subject = sanitize_filename(getattr(attached_item, 'subject', 'Attached Email'))
                    received = getattr(attached_item, 'datetime_received', datetime.now(pytz.UTC))
                    attached_filename = f"attached_email_{attached_subject}_{received.strftime('%Y%m%d%H%M%S')}.html"
                    attached_body = getattr(attached_item, 'body', '')
                    attached_sender = getattr(attached_item, 'sender', None)
                    attached_sender_email = (
                        attached_sender.email_address if getattr(attached_sender, 'email_address', None) else str(attached_sender)
                    )
                    attached_html = (
                        "<html><body\n"
                        f"<h1>Subject: {escape(getattr(attached_item, 'subject', 'No Subject'))}</h1>\n"
                        f"<p><strong>Received:</strong> {escape(received.isoformat())}</p>\n"
                        f"<p><strong>Sender:</strong> {escape(attached_sender_email or 'Unknown')}</p>\n"
                        "<p><strong>Body:</strong></p>\n"
                        f"{attached_body}\n"
                        "</body></html>\n"
                    ).encode('utf-8')
                    session.add(
                        Attachment(
                            email_id=email_record.id,
                            filename=attached_filename,
                            content_type='text/html',
                            data=attached_html,
                            size=len(attached_html),
                        )
                    )
            except Exception as exc:
                logging.error(f"Error saving attachment: {exc}")

    return True


def format_datetime(dt):
    if not dt:
        return 'Unknown'
    try:
        target_tz = pytz.timezone(TIMEZONE)
        if dt.tzinfo is None:
            dt = target_tz.localize(dt)
        else:
            dt = dt.astimezone(target_tz)
        return dt.strftime('%m/%d/%Y %I:%M %p')
    except Exception:
        return dt.strftime('%m/%d/%Y %I:%M %p') if isinstance(dt, datetime) else str(dt)

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
        with session_scope() as session:
            query = (
                session.query(Email)
                .order_by(Email.datetime_received.desc())
                .limit(100)
            )
            recent_emails = []
            for email in query:
                recent_emails.append({
                    'id': email.id,
                    'subject': email.subject or 'No Subject',
                    'sender': email.sender or 'Unknown Sender',
                    'datetime_received': format_datetime(email.datetime_received),
                })

        default_content = "<p class='text-muted text-center'>Select an email to view its contents.</p>"
        return render_template('index.html', emails=recent_emails, email_content=default_content)
    except Exception as e:
        logging.error(f"Error fetching recent emails: {str(e)}")
        default_content = "<p class='text-muted text-center'>Select an email to view its contents.</p>"
        return render_template('index.html', emails=[], email_content=default_content)

@app.route('/search')
@ensure_json_response
def search():
    query = request.args.get('query', '').strip()
    if not query:
        return jsonify({"error": "No search query provided."}), 400

    search_pattern = f"%{query}%"
    lowercase_query = query.lower()
    results = []

    try:
        with session_scope() as session:
            emails = (
                session.query(Email)
                .filter(
                    or_(
                        Email.subject.ilike(search_pattern),
                        Email.sender.ilike(search_pattern),
                        Email.recipients.ilike(search_pattern),
                        Email.body.ilike(search_pattern),
                    )
                )
                .order_by(Email.datetime_received.desc())
                .all()
            )

            for email in emails:
                body_text = email.body or ''
                plain_text = re.sub('<[^<]+?>', '', body_text)
                index = plain_text.lower().find(lowercase_query)
                if index != -1:
                    start = max(0, index - 50)
                    snippet = plain_text[start:index + 150]
                else:
                    snippet = plain_text[:200]
                if snippet:
                    snippet = snippet.strip() + '...'

                results.append({
                    'id': email.id,
                    'subject': email.subject or 'No Subject',
                    'sender': email.sender or 'Unknown Sender',
                    'datetime_received': format_datetime(email.datetime_received),
                    'snippet': snippet,
                })

        return {"results": results}

    except Exception as e:
        logging.error(f"Search error: {str(e)}")
        return {"error": str(e)}, 500

@app.route('/view/<int:email_id>')
def view(email_id):
    try:
        with session_scope() as session:
            email_record = session.get(Email, email_id)
            if not email_record:
                abort(404)
            return build_email_html(email_record)
    except HTTPException:
        raise
    except Exception as e:
        logging.error(f"Error retrieving email {email_id}: {str(e)}")
        abort(500)


@app.route('/attachments/<int:attachment_id>/download')
def download_attachment(attachment_id):
    try:
        with session_scope() as session:
            attachment = session.get(Attachment, attachment_id)
            if not attachment:
                abort(404)
            file_data = attachment.data or b''
            buffer = io.BytesIO(file_data)
            buffer.seek(0)
            return send_file(
                buffer,
                as_attachment=True,
                download_name=attachment.filename or 'attachment',
                mimetype=attachment.content_type or 'application/octet-stream',
            )
    except HTTPException:
        raise
    except Exception as e:
        logging.error(f"Error downloading attachment {attachment_id}: {str(e)}")
        abort(500)


@app.route('/list-attachments/<int:email_id>')
@ensure_json_response
def list_attachments(email_id):
    try:
        with session_scope() as session:
            email_record = session.get(Email, email_id)
            if not email_record:
                return {'attachments': []}

            attachments = []
            for attachment in email_record.attachments:
                attachments.append({
                    'filename': attachment.filename,
                    'path': f"/attachments/{attachment.id}/download",
                    'size': attachment.size or (len(attachment.data) if attachment.data else 0),
                })
            logging.debug(f"Found {len(attachments)} attachments for email {email_id}")
            return {'attachments': attachments}
    except Exception as e:
        logging.error(f"Error listing attachments: {str(e)}")
        return json_response(success=False, message=f"Error listing attachments: {str(e)}", status_code=500)

@app.route('/check-emails', methods=['POST'])
@ensure_json_response
def check_emails():
    try:
        account, output_dir, time_frame = setup_exchange_connection()
        Path(output_dir).mkdir(parents=True, exist_ok=True)

        processed_sent = 0
        processed_inbox = 0

        with session_scope() as session:
            logging.info("Processing sent emails...")
            for created in process_email(account, account.sent, session, time_frame):
                if created:
                    processed_sent += 1

            logging.info("Processing inbox emails...")
            for created in process_email(account, account.inbox, session, time_frame):
                if created:
                    processed_inbox += 1

        return json_response(
            success=True,
            message=f"Emails checked successfully. Processed {processed_sent} sent and {processed_inbox} inbox emails.",
            data={"sent": processed_sent, "inbox": processed_inbox}
        )
    except JSONDecodeError as e:
        logging.error(f"JSON decode error: {str(e)}")
        return json_response(success=False, message=f"JSON decode error: {str(e)}", status_code=500)
    except Exception as e:
        logging.error(f"Failed to check emails: {str(e)}")
        return json_response(success=False, message=f"Failed to check emails: {str(e)}", status_code=500)

@app.route('/download-all-emails')
def download_all_emails():
    """Create and serve a zip file containing all emails and attachments."""
    try:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            with session_scope() as session:
                emails = session.query(Email).options(selectinload(Email.attachments)).all()
                for email_record in emails:
                    recipient_display = (email_record.recipients or 'Unknown').split(',')[0].strip() or 'Unknown'
                    subject_display = sanitize_filename(email_record.subject or 'No_Subject')
                    if email_record.datetime_received:
                        date_str = format_datetime(email_record.datetime_received).replace('/', '-').replace(' ', '_').replace(':', '-')
                    else:
                        date_str = 'Unknown_Date'
                    base_name = sanitize_filename(f"to_{recipient_display} - {subject_display} - {date_str}")

                    zipf.writestr(f"{base_name}.html", build_email_html(email_record))

                    for attachment in email_record.attachments:
                        filename = sanitize_filename(attachment.filename or 'attachment')
                        attachment_path = f"{base_name}_attachments/{filename}"
                        zipf.writestr(attachment_path, attachment.data or b'')

        buffer.seek(0)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        zip_filename = f"all_emails_{timestamp}.zip"
        return send_file(
            buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )

    except Exception as e:
        logging.error(f"Error creating zip file: {str(e)}")
        abort(500)


if __name__ == '__main__':
    # When running the Flask app, you might also want to process emails
    # Uncomment the following lines if you want to process emails on startup
    
    try:
        account, output_dir, time_frame = setup_exchange_connection()
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        with session_scope() as session:
            logging.info("Processing sent emails...")
            for _ in process_email(account, account.sent, session, time_frame):
                pass
            logging.info("Processing inbox emails...")
            for _ in process_email(account, account.inbox, session, time_frame):
                pass
    except Exception as e:
        logging.error(f"Failed to setup Exchange connection or process emails: {str(e)}")
    
    print()
    # app.run(debug=True)
    serve(app, host='127.0.0.1', port=8080)
