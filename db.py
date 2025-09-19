import os
from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker


def _normalize_database_url(url: str) -> str:
    """Ensure the database URL is compatible with SQLAlchemy."""
    if url.startswith("postgres://"):
        return url.replace("postgres://", "postgresql://", 1)
    return url


def get_database_url() -> str:
    """Return the configured database URL or a sensible default."""
    url = os.getenv("DATABASE_URL", "sqlite:///emails.db")
    return _normalize_database_url(url)


database_url = get_database_url()
engine = create_engine(database_url, pool_pre_ping=True, future=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)
Base = declarative_base()


def init_db() -> None:
    """Create database tables if they do not already exist."""
    from models import Email, Attachment  # noqa: F401 - register models

    Base.metadata.create_all(bind=engine)


__all__ = ["SessionLocal", "Base", "init_db", "get_database_url"]
