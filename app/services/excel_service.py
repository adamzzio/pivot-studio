"""Pembacaan workbook XLSX, penyimpanan upload sementara, deteksi tipe kolom, dan ekspor XLSX."""
from __future__ import annotations

import json
import re
import shutil
import threading
import time
import uuid
import zipfile
from collections import OrderedDict
from datetime import date, datetime
from datetime import time as dtime
from decimal import Decimal
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename

from app.utils import AppError

NULL_LABEL = "(kosong)"
VALID_TYPES = ("number", "text", "date")
SID_RE = re.compile(r"^[0-9a-f]{32}$")
EXCEL_MAX_ROWS = 1_048_576

# ---------------------------------------------------------------------------
# Validasi dan inspeksi workbook
# ---------------------------------------------------------------------------


def is_valid_xlsx(src) -> bool:
    try:
        if not zipfile.is_zipfile(src):
            return False
        with zipfile.ZipFile(src) as z:
            return "xl/workbook.xml" in z.namelist()
    except Exception:
        return False


def _safe_int(v):
    try:
        return int(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def inspect_workbook(src) -> list[dict]:
    try:
        wb = load_workbook(src, read_only=True, data_only=True)
    except Exception as exc:
        raise AppError("File XLSX rusak atau tidak dapat dibaca.", 400) from exc
    try:
        sheets = [
            {"name": ws.title, "rows": _safe_int(ws.max_row), "cols": _safe_int(ws.max_column)}
            for ws in wb.worksheets
        ]
    finally:
        wb.close()
    if not sheets:
        raise AppError("Workbook tidak memiliki sheet.", 400)
    return sheets


# ---------------------------------------------------------------------------
# Penyimpanan upload sementara (per sesi)
# ---------------------------------------------------------------------------


def session_dir(upload_dir, sid: str) -> Path:
    if not SID_RE.match(sid or ""):
        raise AppError("Sesi tidak valid.", 400)
    d = Path(upload_dir) / sid
    d.mkdir(parents=True, exist_ok=True)
    return d


def save_upload(storage, upload_dir, sid: str) -> dict:
    display = Path(storage.filename or "workbook.xlsx").name[:200]
    safe = secure_filename(display) or "workbook.xlsx"
    if not safe.lower().endswith(".xlsx"):
        raise AppError(f"'{display}': hanya file .xlsx yang didukung.", 400)
    file_id = uuid.uuid4().hex
    d = session_dir(upload_dir, sid)
    path = d / f"{file_id}.xlsx"
    storage.save(path)
    try:
        if path.stat().st_size == 0:
            raise AppError(f"'{display}': file kosong.", 400)
        if not is_valid_xlsx(path):
            raise AppError(f"'{display}': isi file bukan XLSX yang valid.", 400)
        try:
            sheets = inspect_workbook(path)
        except AppError as exc:
            raise AppError(f"'{display}': {exc.message}", 400) from exc
    except Exception:
        path.unlink(missing_ok=True)
        raise
    meta = {"file_id": file_id, "name": display, "sheets": sheets, "created": time.time()}
    (d / f"{file_id}.json").write_text(json.dumps(meta), encoding="utf-8")
    return meta


def list_files(upload_dir, sid: str) -> list[dict]:
    d = session_dir(upload_dir, sid)
    out = []
    for p in d.glob("*.json"):
        try:
            meta = json.loads(p.read_text(encoding="utf-8"))
            if (d / f"{meta['file_id']}.xlsx").exists():
                out.append(meta)
        except (OSError, ValueError, KeyError):
            continue
    return sorted(out, key=lambda m: m.get("created", 0))


def file_path(upload_dir, sid: str, file_id: str) -> Path:
    if not SID_RE.match(file_id or ""):
        raise AppError("ID file tidak valid.", 400)
    p = session_dir(upload_dir, sid) / f"{file_id}.xlsx"
    if not p.exists():
        raise AppError("File tidak ditemukan atau sesi berakhir. Silakan unggah ulang.", 404)
    return p


def delete_file(upload_dir, sid: str, file_id: str) -> None:
    p = file_path(upload_dir, sid, file_id)
    p.unlink(missing_ok=True)
    p.with_suffix(".json").unlink(missing_ok=True)
    _frames.drop(str(p))


def clear_session(upload_dir, sid: str) -> None:
    d = session_dir(upload_dir, sid)
    shutil.rmtree(d, ignore_errors=True)
    _frames.drop_all()


def cleanup_stale(upload_dir, ttl_seconds: float) -> None:
    root = Path(upload_dir)
    if not root.exists():
        return
    now = time.time()
    for d in root.iterdir():
        if not d.is_dir():
            continue
        try:
            newest = max((f.stat().st_mtime for f in d.iterdir()), default=d.stat().st_mtime)
            if now - newest > ttl_seconds:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            continue


# ---------------------------------------------------------------------------
# Pembacaan sheet dan cache DataFrame
# ---------------------------------------------------------------------------


class _FrameCache:
    def __init__(self, size: int = 6):
        self._d: OrderedDict = OrderedDict()
        self._size = size
        self._lock = threading.Lock()

    def get(self, key):
        with self._lock:
            if key in self._d:
                self._d.move_to_end(key)
                return self._d[key]
        return None

    def put(self, key, value):
        with self._lock:
            self._d[key] = value
            self._d.move_to_end(key)
            while len(self._d) > self._size:
                self._d.popitem(last=False)

    def drop(self, path: str):
        with self._lock:
            for k in [k for k in self._d if k[0] == path]:
                del self._d[k]

    def drop_all(self):
        with self._lock:
            self._d.clear()


_frames = _FrameCache()

_NUM_STD = re.compile(r"^[-+]?\d+(\.\d+)?$")
_NUM_US = re.compile(r"^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$")
_NUM_EU = re.compile(r"^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$")
_NUM_COMMA = re.compile(r"^[-+]?\d+,\d+$")
_LEADING_ZERO = re.compile(r"^[-+]?0\d")


def parse_number_str(s: str, strict: bool = True):
    """Parse angka dari teks. strict=True menolak identifier dengan leading zero."""
    s = s.strip().replace("\u00a0", "")
    if not s:
        return None
    if strict and _LEADING_ZERO.match(s):
        return None
    try:
        if _NUM_STD.match(s):
            return float(s)
        if _NUM_US.match(s):
            return float(s.replace(",", ""))
        if _NUM_EU.match(s):
            return float(s.replace(".", "").replace(",", "."))
        if _NUM_COMMA.match(s):
            return float(s.replace(",", "."))
    except ValueError:
        return None
    return None


def infer_type(series: pd.Series) -> str:
    s = series.dropna()
    if s.empty:
        return "text"
    kind = pd.api.types.infer_dtype(s, skipna=True)
    if kind in ("integer", "floating", "mixed-integer-float", "decimal"):
        return "number"
    if kind in ("datetime", "datetime64", "date"):
        return "date"
    if kind == "string":
        for v in s.iloc[:200]:
            if parse_number_str(v) is None:
                return "text"
        return "number" if all(parse_number_str(v) is not None for v in s) else "text"
    return "text"


def _prepare_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(axis=0, how="all")
    keep = [c for c in df.columns if not (str(c).startswith("Unnamed:") and df[c].isna().all())]
    df = df[keep].copy()
    names, seen = [], set()
    for i, c in enumerate(df.columns, 1):
        n = str(c).strip()
        if not n or re.match(r"^Unnamed: \d+", n):
            n = f"Kolom_{i}"
        base, k = n, 1
        while n in seen:
            k += 1
            n = f"{base}_{k}"
        seen.add(n)
        names.append(n)
    df.columns = names
    for c in df.columns:
        s = df[c]
        if s.dtype == object:
            mask = s.map(lambda v: isinstance(v, str) and not v.strip())
            if mask.any():
                df[c] = s.mask(mask, np.nan)
    df = df.reset_index(drop=True)
    if df.empty or len(df.columns) == 0:
        raise AppError("Sheet kosong atau tidak memiliki data di bawah baris header.", 422)
    return df


def load_frame(path: Path, sheet: str):
    key = (str(path), path.stat().st_mtime_ns, sheet)
    hit = _frames.get(key)
    if hit is not None:
        return hit
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype=object)
    except ValueError as exc:
        raise AppError(f"Sheet '{sheet}' tidak ditemukan di workbook.", 404) from exc
    except Exception as exc:
        raise AppError("Gagal membaca sheet. File mungkin rusak.", 400) from exc
    df = _prepare_frame(df)
    inferred = {c: infer_type(df[c]) for c in df.columns}
    _frames.put(key, (df, inferred))
    return df, inferred


# ---------------------------------------------------------------------------
# Konversi tipe, label, dan format tampilan
# ---------------------------------------------------------------------------


def fmt_float(v: float) -> str:
    if v != v:
        return ""
    if float(v).is_integer() and abs(v) < 1e15:
        return str(int(v))
    return format(v, ".15g")


def fmt_datetime(v) -> str:
    if v.hour == 0 and v.minute == 0 and v.second == 0:
        return v.strftime("%Y-%m-%d")
    return v.strftime("%Y-%m-%d %H:%M:%S")


def text_value(v):
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    try:
        if pd.isna(v):
            return np.nan
    except (TypeError, ValueError):
        pass
    if isinstance(v, (datetime, pd.Timestamp)):
        return fmt_datetime(v)
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, dtime):
        return v.isoformat()
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating, Decimal)):
        return fmt_float(float(v))
    return str(v)


def display_value(v) -> str:
    out = text_value(v)
    return "" if isinstance(out, float) else out


def _to_float(v):
    if isinstance(v, bool):
        return np.nan
    if isinstance(v, (int, float, np.number, Decimal)):
        return float(v)
    if isinstance(v, str):
        r = parse_number_str(v, strict=False)
        return np.nan if r is None else r
    return np.nan


def _to_number(s: pd.Series):
    kind = pd.api.types.infer_dtype(s.dropna(), skipna=True)
    if kind in ("integer", "floating", "mixed-integer-float"):
        out = pd.to_numeric(s, errors="coerce").astype("float64")
    else:
        out = s.map(_to_float).astype("float64")
    bad = int((s.notna() & out.isna()).sum())
    return out, bad


def _to_date(s: pd.Series):
    out = pd.to_datetime(s, errors="coerce", dayfirst=True, format="mixed")
    bad = int((s.notna() & out.isna()).sum())
    return out, bad


def apply_types(df: pd.DataFrame, inferred: dict, overrides: dict, fields) -> tuple:
    """Kembalikan (DataFrame bertipe, peta tipe, catatan) hanya untuk kolom yang dibutuhkan."""
    typed, types, notes = {}, {}, []
    for col in dict.fromkeys(fields):
        ov = overrides.get(col)
        t = ov if ov in VALID_TYPES else inferred[col]
        s = df[col]
        if t == "number":
            ser, bad = _to_number(s)
            if bad:
                notes.append(f"{bad} nilai pada kolom '{col}' bukan angka dan diabaikan.")
        elif t == "date":
            ser, bad = _to_date(s)
            if bad:
                notes.append(f"{bad} nilai pada kolom '{col}' bukan tanggal dan diabaikan.")
        else:
            ser = s.map(text_value)
        typed[col], types[col] = ser, t
    return pd.DataFrame(typed, index=df.index), types, notes


def label_series(s: pd.Series, t: str) -> pd.Series:
    if t == "number":
        m = {u: fmt_float(float(u)) for u in s.dropna().unique()}
        out = s.map(m)
    elif t == "date":
        nn = s.dropna()
        has_time = bool(len(nn)) and bool(
            ((nn.dt.hour != 0) | (nn.dt.minute != 0) | (nn.dt.second != 0)).any()
        )
        out = s.dt.strftime("%Y-%m-%d %H:%M:%S" if has_time else "%Y-%m-%d")
    else:
        out = s
    return out.fillna(NULL_LABEL).astype(str)


# ---------------------------------------------------------------------------
# Ekspor XLSX
# ---------------------------------------------------------------------------

_NAVY = PatternFill("solid", fgColor="1E3A8A")
_NAVY_SOFT = PatternFill("solid", fgColor="DBEAFE")
_NAVY_MID = PatternFill("solid", fgColor="BFDBFE")
_THIN = Side(style="thin", color="CBD5E1")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_INT_AGGS = {"COUNT", "COUNT_NON_NULL", "COUNT_DISTINCT"}


def _clean_str(v: str) -> str:
    return ILLEGAL_CHARACTERS_RE.sub("", v) if ILLEGAL_CHARACTERS_RE.search(v) else v


def _put(ws, r: int, c: int, v):
    """Tulis nilai; teks yang diawali '=' tidak boleh dianggap formula."""
    if isinstance(v, str):
        v = _clean_str(v)
        cell = ws.cell(row=r, column=c, value=v)
        if v.startswith("="):
            cell.data_type = "s"
        return cell
    return ws.cell(row=r, column=c, value=v)


def _num_format(agg: str, decimals: int) -> str:
    if agg in _INT_AGGS or decimals <= 0:
        return "#,##0"
    return "#,##0." + "0" * decimals


def write_pivot_workbook(result: dict, raw_df: pd.DataFrame, filters_text: list[str], decimals: int) -> Workbook:
    decimals = max(0, min(int(decimals), 8))
    wb = Workbook()
    ws = wb.active
    ws.title = "Pivot"
    hdr, meta = result["header"], result["meta"]
    metrics, M = hdr["metrics"], len(hdr["metrics"])
    row_fields, col_fields = hdr["row_fields"], hdr["col_fields"]
    col_paths, grand_col = hdr["col_paths"], hdr["grand_col"]
    nrc, C = max(1, len(row_fields)), len(col_fields)

    raw_limit = EXCEL_MAX_ROWS - 1
    raw_truncated = len(raw_df) > raw_limit

    info = [
        "Pivot Table",
        "Filter aktif: " + ("; ".join(filters_text) if filters_text else "tidak ada"),
        f"Record dianalisis: {meta['records']} dari {meta['total_records']}",
        "Nilai: " + ", ".join(f"{m['label']} ({m['agg']} pada {m['field']})" for m in metrics),
    ]
    if raw_truncated:
        info.append(f"Catatan: sheet Raw Data dipotong pada {raw_limit} baris (batas Excel).")
    for i, line in enumerate(info, 1):
        _put(ws, i, 1, line)
    ws["A1"].font = Font(bold=True, size=14)
    start = len(info) + 2

    white_bold = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def head(r, c, v, span_c=1, span_r=1):
        cell = _put(ws, r, c, v)
        cell.font, cell.fill, cell.alignment = white_bold, _NAVY, center
        if span_c > 1 or span_r > 1:
            ws.merge_cells(start_row=r, start_column=c, end_row=r + span_r - 1, end_column=c + span_c - 1)
        for rr in range(r, r + span_r):
            for cc in range(c, c + span_c):
                ws.cell(row=rr, column=cc).fill = _NAVY
                ws.cell(row=rr, column=cc).border = _BORDER

    if C == 0:
        head_rows = 1
        for i in range(nrc):
            head(start, 1 + i, row_fields[i] if row_fields else "Total")
        for j, m in enumerate(metrics):
            head(start, nrc + 1 + j, m["label"])
    else:
        head_rows = C + (1 if M > 1 else 0)
        for i in range(nrc):
            head(start, 1 + i, row_fields[i] if row_fields else "Total", 1, head_rows)
        for lvl in range(C):
            col, i = nrc + 1, 0
            while i < len(col_paths):
                j = i
                while j < len(col_paths) and col_paths[j][: lvl + 1] == col_paths[i][: lvl + 1]:
                    j += 1
                span = (j - i) * M
                head(start + lvl, col, col_paths[i][lvl], span)
                col += span
                i = j
        if grand_col:
            head(start, nrc + 1 + len(col_paths) * M, "Grand Total", M, C)
        if M > 1:
            blocks = len(col_paths) + (1 if grand_col else 0)
            for b in range(blocks):
                for j, m in enumerate(metrics):
                    head(start + C, nrc + 1 + b * M + j, m["label"])

    r = start + head_rows
    for entry in result["body"]:
        kind, path = entry["kind"], entry["labels"]
        labels = [""] * nrc
        if kind == "grand":
            labels[0] = "Grand Total"
        else:
            for i, p in enumerate(path[:nrc]):
                labels[i] = p
            if kind == "subtotal" and len(path) < nrc:
                labels[len(path)] = "Subtotal"
        for i, lab in enumerate(labels):
            _put(ws, r, 1 + i, lab).border = _BORDER
        for j, v in enumerate(entry["values"]):
            agg = metrics[j % M]["agg"]
            cell = _put(ws, r, nrc + 1 + j, v)
            cell.border = _BORDER
            if isinstance(v, (int, float)):
                cell.number_format = _num_format(agg, decimals)
                cell.alignment = Alignment(horizontal="right")
        if kind != "detail":
            fill = _NAVY_MID if kind == "grand" else _NAVY_SOFT
            for cc in range(1, nrc + len(entry["values"]) + 1):
                ws.cell(row=r, column=cc).fill = fill
                ws.cell(row=r, column=cc).font = Font(bold=True)
        r += 1

    ws.freeze_panes = ws.cell(row=start + head_rows, column=nrc + 1)
    for c in range(1, nrc + 1):
        ws.column_dimensions[get_column_letter(c)].width = 22
    for c in range(nrc + 1, nrc + 1 + (len(col_paths) + 1) * M + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16

    # Sheet raw data (nilai asli, tanpa konversi tipe)
    rs = wb.create_sheet("Raw Data")
    for c, name in enumerate(raw_df.columns, 1):
        cell = _put(rs, 1, c, str(name))
        cell.font, cell.fill = white_bold, _NAVY
    sub = raw_df.iloc[:raw_limit]
    for ri, row in enumerate(sub.itertuples(index=False, name=None), 2):
        for ci, v in enumerate(row, 1):
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                continue
            if isinstance(v, str):
                _put(rs, ri, ci, v)
            else:
                rs.cell(row=ri, column=ci, value=v)
    rs.freeze_panes = "A2"
    return wb
