import json
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

from flask import Blueprint, render_template, request, send_file
from werkzeug.utils import secure_filename

from app.services import conversion_service as cs
from app.utils import AppError, ok

bp = Blueprint("converter", __name__)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _read_upload(ext: str) -> tuple[bytes, str]:
    f = request.files.get("file")
    if f is None or not f.filename:
        raise AppError("Pilih file terlebih dahulu.", 400)
    name = Path(f.filename).name
    if not name.lower().endswith(ext):
        raise AppError(f"Format file tidak sesuai. Unggah file {ext}.", 400)
    raw = f.read()
    if not raw:
        raise AppError("File kosong.", 400)
    return raw, secure_filename(Path(name).stem) or "converted"


@bp.get("/converter")
def page():
    return render_template("converter.html")


@bp.post("/api/convert/inspect")
def inspect():
    kind = request.form.get("kind")
    if kind == "csv-to-xlsx":
        raw, _ = _read_upload(".csv")
        return ok(info=cs.inspect_csv(raw))
    if kind == "xlsx-to-csv":
        raw, _ = _read_upload(".xlsx")
        return ok(info=cs.inspect_xlsx(raw))
    raise AppError("Jenis konversi tidak dikenali.", 400)


@bp.post("/api/convert/csv-to-xlsx")
def csv_to_xlsx():
    raw, stem = _read_upload(".csv")
    data, notes = cs.csv_to_xlsx(
        raw,
        delimiter=request.form.get("delimiter", "auto"),
        encoding=request.form.get("encoding", "auto"),
        type_mode=request.form.get("type_mode", "text"),
    )
    resp = send_file(BytesIO(data), mimetype=XLSX_MIME, as_attachment=True, download_name=f"{stem}.xlsx")
    resp.headers["X-Notes"] = quote(json.dumps(notes))
    return resp


@bp.post("/api/convert/xlsx-to-csv")
def xlsx_to_csv():
    raw, stem = _read_upload(".xlsx")
    data, notes = cs.xlsx_to_csv(
        raw,
        sheet=request.form.get("sheet", ""),
        delimiter=request.form.get("delimiter", "comma"),
        encoding=request.form.get("encoding", "utf-8-sig"),
    )
    resp = send_file(BytesIO(data), mimetype="text/csv", as_attachment=True, download_name=f"{stem}.csv")
    resp.headers["X-Notes"] = quote(json.dumps(notes))
    return resp
