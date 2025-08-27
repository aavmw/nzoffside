# backend/tests/conftest.py
import os
import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

# Import your app factory and models
from app.wsgi import create_app
from app import db as dbmod
from app.models import Base


@pytest.fixture(scope="session")
def test_db_url():
    # Prefer a dedicated TEST DB if you have it; otherwise use your dev DB.
    return os.environ.get(
        "DATABASE_URL", "postgresql+psycopg://postgres:546335@localhost:5433/dwh_local"
    )


@pytest.fixture(scope="session")
def engine(test_db_url):
    eng = create_engine(test_db_url, pool_pre_ping=True, future=True)
    # Create schema for tests
    Base.metadata.create_all(eng)
    yield eng
    # Drop schema after tests
    Base.metadata.drop_all(eng)


@pytest.fixture()
def app(engine, monkeypatch):
    # Rewire the global SessionLocal/engine used by the app to the test engine
    dbmod.engine = engine
    dbmod.SessionLocal = sessionmaker(
        bind=engine, autoflush=False, autocommit=False, future=True
    )

    # Make sure API key passes in tests
    os.environ["API_KEY_DEV"] = "test-key"

    flask_app = create_app()
    flask_app.config.update(TESTING=True)
    return flask_app


@pytest.fixture()
def client(app):
    return app.test_client()


@pytest.fixture()
def auth_header():
    return {"X-Api-Key": "test-key"}
