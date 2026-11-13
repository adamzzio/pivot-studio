import io
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import create_app  # noqa: E402
from config import Config  # noqa: E402


@pytest.fixture()
def app(tmp_path):
    class TestConfig(Config):
        TESTING = True
        UPLOAD_DIR = tmp_path / "uploads"
        INSTANCE_DIR = tmp_path
        GROQ_API_KEY = ""

    return create_app(TestConfig)


@pytest.fixture()
def client(app):
    return app.test_client()


def make_xlsx(rows, sheet="Data") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


SALES = [
    ["Region", "Product", "Sales", "Qty", "Code"],
    ["A", "X", 10, 1, "001"],
    ["A", "Y", 20, 2, "002"],
    ["A", "X", 30, 2, "003"],
    ["B", "X", 40, 4, "004"],
    ["B", "Y", None, 4, "005"],
    ["B", "Y", 60, 5, "006"],
]
