"""Konversi CSV <-> XLSX tanpa mengubah nilai secara diam-diam."""
from __future__ import annotations

import csv
import io
import re
from datetime import date, datetime
from datetime import time as dtime

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font

from app.utils import AppError

from . import excel_service as xs

DELIMS = {"comma": ",", "semicolon": ";", "tab": "\t", "pipe": "|"}
CSV_ENCODINGS = {"utf-8": "utf-8", "utf-8-sig": "utf-8-sig", "windows-1252": "cp1252"}
_PLAIN_NUM = r"^[-+]?(?:0|[1-9]\d*)(?:\.\d+)?$"
_NUMERIC_LIKE = r"^[-+]?\d+(?:\.\d+)?$"
_LEAD_ZERO = r"^[-+]?0\d"


def delimiter_name(d: str) -> str:
    return {",": "koma (,)", ";": "titik koma (;)", "\t": "tab", "|": "pipa (|)"}.get(d, d)


def decode_csv(raw: bytes, encoding: str = "auto") -> tuple[str, str]:
    if not raw.strip():
        raise AppError("File CSV kosong.", 400)
    if encoding != "auto":
        enc = CSV_ENCODINGS.get(encoding)
        if not enc:
            raise AppError("Encoding tidak didukung.", 400)
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError as exc:
            raise AppError("File tidak dapat dibaca dengan encoding yang dipilih. Coba encoding lain.", 400) from exc
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise AppError("Encoding file tidak dikenali.", 400)


def detect_delimiter(text: str) -> str:
    lines = text[:65536].splitlines()[:5]
    best, score = ",", 0
    for d in (",", ";", "\t", "|"):
        n = sum(line.count(d) for line in lines)
        if n > score:
            best, score = d, n
    return best


def inspect_csv(raw: bytes) -> dict:
    text, enc = decode_csv(raw)
    d = detect_delimiter(text)
    return {"kind": "csv", "delimiter": d, "delimiter_name": delimiter_name(d), "encoding": enc}


def inspect_xlsx(raw: bytes) -> dict:
    bio = io.BytesIO(raw)
    if not xs.is_valid_xlsx(bio):
        raise AppError("File bukan XLSX yang valid atau rusak.", 400)
    bio.seek(0)
    return {"kind": "xlsx", "sheets": xs.inspect_workbook(bio)}


def _col_plan(df: pd.DataFrame, mode: str) -> tuple[dict, list[str]]:
    plan, notes = {}, []
    for c in df.columns:
        s = df[c]
        ne = s[s != ""]
        if mode != "auto" or ne.empty:
            plan[c] = "text"
            continue
        if ne.str.fullmatch(_PLAIN_NUM).all():
            digits = ne.str.replace(r"\D", "", regex=True).str.len().max()
            if digits <= 15:
                plan[c] = "int" if not ne.str.contains(r"\.", regex=True).any() else "float"
                continue
            notes.append(f"Kolom '{c}' dipertahankan sebagai teks karena angkanya lebih dari 15 digit.")
        elif ne.str.fullmatch(_NUMERIC_LIKE).all() and ne.str.match(_LEAD_ZERO).any():
            notes.append(f"Kolom '{c}' dipertahankan sebagai teks karena memiliki leading zero.")
        plan[c] = "text"
    return plan, notes


def _text_cell(ws, v: str):
    if ILLEGAL_CHARACTERS_RE.search(v):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)
    if v.startswith("="):
        cell = WriteOnlyCell(ws, value=v)
        cell.data_type = "s"
        return cell
    return v


def csv_to_xlsx(raw: bytes, delimiter: str = "auto", encoding: str = "auto", type_mode: str = "text") -> tuple[bytes, list[str]]:
    text, enc = decode_csv(raw, encoding)
    delim = detect_delimiter(text) if delimiter == "auto" else DELIMS.get(delimiter)
    if not delim:
        raise AppError("Delimiter tidak didukung.", 400)
    try:
        df = pd.read_csv(io.StringIO(text), sep=delim, dtype=str, keep_default_na=False, na_filter=False)
    except pd.errors.EmptyDataError as exc:
        raise AppError("File CSV kosong atau tidak memiliki header.", 400) from exc
    except pd.errors.ParserError as exc:
        raise AppError("CSV tidak valid: jumlah kolom antar-baris tidak konsisten. Coba ganti delimiter.", 400) from exc
    if len(df) + 1 > xs.EXCEL_MAX_ROWS:
        raise AppError("Jumlah baris melebihi batas Excel (1.048.576 baris).", 400)

    cols = []
    for i, c in enumerate(df.columns, 1):
        n = str(c).strip()
        cols.append(f"Kolom_{i}" if (not n or re.match(r"^Unnamed: \d+$", n)) else n)
    df.columns = cols

    plan, notes = _col_plan(df, type_mode)
    notes.insert(0, f"Terbaca sebagai CSV dengan delimiter {delimiter_name(delim)} dan encoding {enc}.")

    lists = []
    for c in df.columns:
        s = df[c]
        kind = plan[c]
        if kind == "text":
            lists.append(None)
        else:
            num = pd.to_numeric(s.where(s != ""), errors="coerce")
            if kind == "int":
                num = num.astype("Int64")
            lists.append([None if pd.isna(x) else x for x in num.tolist()])

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    ws.freeze_panes = "A2"
    header = []
    for c in df.columns:
        cell = WriteOnlyCell(ws, value=_clean_header(c))
        cell.font = Font(bold=True)
        header.append(cell)
    ws.append(header)
    text_cols = [df[c].tolist() if plan[c] == "text" else None for c in df.columns]
    for i in range(len(df)):
        row = []
        for j, c in enumerate(df.columns):
            if lists[j] is not None:
                row.append(lists[j][i])
            else:
                v = text_cols[j][i]
                row.append(None if v == "" else _text_cell(ws, v))
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), notes


def _clean_header(v: str) -> str:
    return ILLEGAL_CHARACTERS_RE.sub("", v)


def _fmt_cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second, v.microsecond) == (0, 0, 0, 0) else v.isoformat(sep=" ")
    if isinstance(v, (date, dtime)):
        return v.isoformat()
    if isinstance(v, float):
        return xs.fmt_float(v)
    return str(v)


def xlsx_to_csv(raw: bytes, sheet: str, delimiter: str = "comma", encoding: str = "utf-8-sig") -> tuple[bytes, list[str]]:
    delim = DELIMS.get(delimiter)
    if not delim:
        raise AppError("Delimiter tidak didukung.", 400)
    if encoding not in ("utf-8", "utf-8-sig"):
        raise AppError("Encoding output tidak didukung.", 400)
    bio = io.BytesIO(raw)
    if not xs.is_valid_xlsx(bio):
        raise AppError("File bukan XLSX yang valid atau rusak.", 400)
    bio.seek(0)
    try:
        wb = load_workbook(bio, read_only=True, data_only=True)
    except Exception as exc:
        raise AppError("File XLSX rusak atau tidak dapat dibaca.", 400) from exc
    try:
        if sheet not in wb.sheetnames:
            raise AppError("Sheet yang dipilih tidak ditemukan.", 400)
        rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
    finally:
        wb.close()
    while rows and all(v is None or v == "" for v in rows[-1]):
        rows.pop()
    if not rows:
        raise AppError("Sheet yang dipilih kosong.", 400)
    width = 0
    for r in rows:
        for i in range(len(r) - 1, -1, -1):
            if r[i] is not None and r[i] != "":
                width = max(width, i + 1)
                break
    out = io.StringIO(newline="")
    w = csv.writer(out, delimiter=delim, lineterminator="\r\n")
    for r in rows:
        r = (r + [None] * width)[:width]
        w.writerow([_fmt_cell(v) for v in r])
    notes = [
        f"Sheet '{sheet}' diekspor ({len(rows)} baris, {width} kolom). Formatting, formula, dan fitur khusus XLSX tidak dipertahankan di CSV."
    ]
    return out.getvalue().encode(encoding), notes
