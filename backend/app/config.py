# backend/app/config.py
from functools import lru_cache
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict

# Load your dev env early (optional)
load_dotenv(Path(__file__).parents[1] / ".env.dev", override=False)

class Settings(BaseSettings):
    # Config for pydantic-settings
    model_config = SettingsConfigDict(
        env_file=".env",            # used if present; real env vars still take precedence
        env_file_encoding="utf-8",
        case_sensitive=True,
        extra="ignore",
    )

    # --- Core ---
    ENV: str = Field(default="development")
    DEBUG: bool = Field(default=False)
    SECRET_KEY: str = Field(default="dev-secret")

    # --- DB ---
    # Keep it simple as str; you can switch to typed DSN later
    DATABASE_URL: str

    # --- API auth ---
    API_KEY_DEV: Optional[str] = None

    # --- Google / Sheets ---
    GOOGLE_CREDS_PATH: str = "./google_creds.json"
    MASTER_SPREADSHEET_ID: str = ""

    # --- Misc ---
    TZ: str = "Europe/Prague"

@lru_cache
def get_settings() -> Settings:
    return Settings()
