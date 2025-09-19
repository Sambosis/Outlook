import os
import psycopg2
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker
from sqlalchemy.exc import OperationalError
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def get_db_connection():
    """Establishes a connection to the database using the DATABASE_URL environment variable."""
    db_url = os.getenv('DATABASE_URL')
    if not db_url:
        raise ValueError("DATABASE_URL environment variable is not set.")

    try:
        # Replace postgres:// with postgresql:// for SQLAlchemy compatibility
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql://", 1)

        engine = create_engine(db_url)
        conn = engine.connect()
        return conn
    except OperationalError as e:
        logger.error(f"Could not connect to the database: {e}")
        raise

def create_tables():
    """Creates the emails and attachments tables if they don't already exist."""
    conn = None
    try:
        conn = get_db_connection()
        create_emails_table = text("""
        CREATE TABLE IF NOT EXISTS emails (
            id SERIAL PRIMARY KEY,
            subject TEXT,
            sender TEXT,
            recipients TEXT,
            body TEXT,
            received_at TIMESTAMP WITH TIME ZONE,
            created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );
        """)
        create_attachments_table = text("""
        CREATE TABLE IF NOT EXISTS attachments (
            id SERIAL PRIMARY KEY,
            email_id INTEGER REFERENCES emails(id) ON DELETE CASCADE,
            filename TEXT NOT NULL,
            content BYTEA NOT NULL,
            created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );
        """)
        conn.execute(create_emails_table)
        conn.execute(create_attachments_table)
        logger.info("Tables 'emails' and 'attachments' are ready.")
    finally:
        if conn:
            conn.close()

if __name__ == '__main__':
    # This allows running this script directly to set up the database tables.
    logger.info("Setting up database tables...")
    create_tables()
    logger.info("Database setup complete.")
