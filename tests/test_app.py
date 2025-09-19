# tests/test_app.py

import os
import sys
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

os.environ.setdefault('DATABASE_URL', 'sqlite:///test.db')

from app import app

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
