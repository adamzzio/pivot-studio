"""Konfigurasi aplikasi. Semua secret dibaca dari environment variable."""
import os
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


def _float(name: str, default: float) -> float:
    try:
        return float(os.getenv(name, default))
    except ValueError:
        return default


class Config:
    SECRET_KEY = os.getenv("SECRET_KEY") or os.urandom(24).hex()
    INSTANCE_DIR = BASE_DIR / "instance"
    UPLOAD_DIR = INSTANCE_DIR / "uploads"

    # Tanpa batas ukuran upload (None). Isi MAX_CONTENT_MB di .env bila ingin membatasi.
    MAX_CONTENT_LENGTH = (
        int(float(os.environ["MAX_CONTENT_MB"]) * 1024 * 1024)
        if os.getenv("MAX_CONTENT_MB")
        else None
    )
    FILE_TTL_HOURS = _float("FILE_TTL_HOURS", 12)

    SESSION_COOKIE_SAMESITE = "Lax"
    SESSION_COOKIE_HTTPONLY = True

    # Pivot
    PIVOT_PAGE_SIZE = 500
    PIVOT_MAX_CELLS = 3_000_000

    # Groq
    GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
    GROQ_MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-120b")
    GROQ_TIMEOUT = _float("GROQ_TIMEOUT", 60)
    AI_MAX_ROWS = 150
    AI_MAX_COLS = 24
