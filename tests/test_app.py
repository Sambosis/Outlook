# tests/test_app.py

import os
import sys
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from sqlalchemy import create_engine, inspect, text

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

os.environ.setdefault('DATABASE_URL', 'sqlite:///test.db')

from app import app, ensure_database_schema

@pytest.fixture
def client():
    with app.test_client() as client:
        yield client

def test_index(client):
    response = client.get('/')
    assert response.status_code == 200
    assert b'Email Search' in response.data

def test_check_emails(client):
    with patch('app.setup_exchange_connection') as mock_setup, patch('app.process_email') as mock_process:
        mock_setup.return_value = (MagicMock(), 'output', datetime.utcnow())
        mock_process.return_value = []
        response = client.post('/check-emails')
    assert response.status_code == 200
    assert b'success' in response.data

def test_search(client):
    response = client.get('/search?query=test')
    assert response.status_code == 200
    assert b'results' in response.data

def test_view(client):
    response = client.get('/view/1')
    assert response.status_code == 200 or response.status_code == 404

def test_download_attachment(client):
    response = client.get('/attachments/1/download')
    assert response.status_code == 200 or response.status_code == 404

def test_list_attachments(client):
    response = client.get('/list-attachments/1')
    assert response.status_code == 200
    assert b'attachments' in response.data


def test_ensure_database_schema_adds_created_at_columns(tmp_path):
    test_db_path = tmp_path / 'legacy.db'
    engine = create_engine(f'sqlite:///{test_db_path}')

    with engine.begin() as connection:
        connection.execute(
            text(
                "CREATE TABLE emails (\n"
                "    id INTEGER PRIMARY KEY,\n"
                "    subject TEXT\n"
                ")"
            )
        )
        connection.execute(
            text(
                "CREATE TABLE attachments (\n"
                "    id INTEGER PRIMARY KEY,\n"
                "    email_id INTEGER,\n"
                "    filename TEXT\n"
                ")"
            )
        )
        connection.execute(
            text(
                "INSERT INTO attachments (email_id, filename) VALUES (1, 'test.txt')"
            )
        )

    ensure_database_schema(engine)

    inspector = inspect(engine)
    email_columns = {column['name'] for column in inspector.get_columns('emails')}
    attachment_columns = {column['name'] for column in inspector.get_columns('attachments')}

    assert 'created_at' in email_columns
    assert 'created_at' in attachment_columns

    with engine.connect() as connection:
        created_at_value = connection.execute(
            text("SELECT created_at FROM attachments LIMIT 1")
        ).scalar()

    assert created_at_value is not None
