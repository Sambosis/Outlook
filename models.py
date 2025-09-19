from datetime import datetime
from sqlalchemy import Column, DateTime, ForeignKey, Integer, LargeBinary, String, Text
from sqlalchemy.orm import relationship

from db import Base


class Email(Base):
    __tablename__ = "emails"

    id = Column(Integer, primary_key=True)
    message_id = Column(String(512), unique=True, index=True, nullable=True)
    subject = Column(String(1024), nullable=True)
    sender = Column(String(512), nullable=True)
    recipients = Column(Text, nullable=True)
    primary_recipient = Column(String(512), nullable=True)
    folder = Column(String(64), nullable=True)
    datetime_received = Column(DateTime, nullable=True)
    body = Column(Text, nullable=True)
    body_plain = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    attachments = relationship(
        "Attachment", back_populates="email", cascade="all, delete-orphan", lazy="selectin"
    )


class Attachment(Base):
    __tablename__ = "attachments"

    id = Column(Integer, primary_key=True)
    email_id = Column(Integer, ForeignKey("emails.id", ondelete="CASCADE"), nullable=False, index=True)
    filename = Column(String(512), nullable=True)
    content_type = Column(String(255), nullable=True)
    size = Column(Integer, nullable=True)
    content_id = Column(String(255), nullable=True)
    data = Column(LargeBinary, nullable=True)

    email = relationship("Email", back_populates="attachments")


__all__ = ["Email", "Attachment"]
