import os
import sys
from datetime import datetime

import pytest

sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from app import app
from db import SessionLocal
from models import Attachment, Email


@pytest.fixture
def client():
    with app.test_client() as client:
        yield client


@pytest.fixture
def sample_email():
    with SessionLocal() as session:
        email = Email(
            subject='Test Email',
            sender='sender@example.com',
            recipients='recipient@example.com',
            primary_recipient='recipient@example.com',
            datetime_received=datetime.utcnow(),
            body='<p>Hello World</p>',
            body_plain='Hello World',
        )
        session.add(email)
        session.commit()
        session.refresh(email)

        attachment = Attachment(
            email_id=email.id,
            filename='test.txt',
            content_type='text/plain',
            size=4,
            data=b'test'
        )
        session.add(attachment)
        session.commit()
        session.refresh(attachment)

        yield email, attachment

        session.query(Attachment).filter_by(email_id=email.id).delete()
        session.query(Email).filter_by(id=email.id).delete()
        session.commit()


def test_index(client):
    response = client.get('/')
    assert response.status_code == 200
    assert b'Email Search' in response.data


def test_search(client):
    response = client.get('/search?query=test')
    assert response.status_code == 200
    assert b'results' in response.data


def test_view_email(client, sample_email):
    email, _ = sample_email
    response = client.get(f'/email/{email.id}')
    assert response.status_code == 200
    assert b'Test Email' in response.data


def test_download_attachment(client, sample_email):
    _, attachment = sample_email
    response = client.get(f'/attachments/{attachment.id}/download')
    assert response.status_code == 200
    assert response.data == b'test'


def test_list_attachments(client, sample_email):
    email, attachment = sample_email
    response = client.get(f'/attachments/{email.id}')
    assert response.status_code == 200
    json_data = response.get_json()
    assert 'attachments' in json_data
    assert any(att['id'] == attachment.id for att in json_data['attachments'])


@pytest.mark.skip(reason="Requires live Exchange connection")
def test_check_emails(client):
    response = client.post('/check-emails')
    assert response.status_code in (200, 500)
