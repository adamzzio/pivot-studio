from io import BytesIO

from flask import Blueprint, current_app, render_template, request, send_file, session

from app.services import excel_service as xs
from app.services import pivot_service as ps
from app.utils import AppError, ok

bp = Blueprint("pivot", __name__)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _sid() -> str:
    return session["sid"]


def _updir():
    return current_app.config["UPLOAD_DIR"]


def _json() -> dict:
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        raise AppError("Body permintaan harus berupa JSON.", 400)
    return data


def _types(data: dict) -> dict:
    t = data.get("types")
    return {str(k): v for k, v in t.items() if v in xs.VALID_TYPES} if isinstance(t, dict) else {}


def _frame(data: dict):
    file_id, sheet = data.get("file_id"), data.get("sheet")
    if not isinstance(file_id, str) or not isinstance(sheet, str) or not sheet:
        raise AppError("Sumber data belum dipilih.", 400)
    path = xs.file_path(_updir(), _sid(), file_id)
    return xs.load_frame(path, sheet)


@bp.get("/pivot")
def page():
    return render_template("pivot.html")


@bp.get("/api/pivot/files")
def files():
    return ok(files=xs.list_files(_updir(), _sid()))


@bp.post("/api/pivot/upload")
def upload():
    items = request.files.getlist("files")
    if not items:
        raise AppError("Pilih minimal satu file XLSX.", 400)
    saved, errors = [], []
    for item in items:
        try:
            saved.append(xs.save_upload(item, _updir(), _sid()))
        except AppError as exc:
            errors.append({"name": (item.filename or "file"), "error": exc.message})
    if not saved:
        raise AppError(errors[0]["error"] if len(errors) == 1 else "Tidak ada file yang berhasil diunggah.", 400)
    return ok(files=saved, errors=errors)


@bp.delete("/api/pivot/file/<file_id>")
def delete(file_id: str):
    xs.delete_file(_updir(), _sid(), file_id)
    return ok()


@bp.post("/api/pivot/clear")
def clear():
    xs.clear_session(_updir(), _sid())
    return ok()


@bp.post("/api/pivot/preview")
def preview():
    data = _json()
    df, inferred = _frame(data)
    overrides = _types(data)
    try:
        offset = max(0, int(data.get("offset", 0)))
        limit = min(200, max(1, int(data.get("limit", 25))))
    except (TypeError, ValueError):
        raise AppError("Parameter halaman tidak valid.", 400)
    page_df = df.iloc[offset : offset + limit]
    columns = [
        {
            "name": c,
            "type": overrides.get(c) or inferred[c],
            "inferred": inferred[c],
            "nulls": int(df[c].isna().sum()),
        }
        for c in df.columns
    ]
    rows = [[xs.display_value(v) for v in r] for r in page_df.itertuples(index=False, name=None)]
    return ok(columns=columns, rows=rows, total_rows=int(len(df)), total_cols=int(len(df.columns)), offset=offset)


@bp.post("/api/pivot/field-values")
def field_values():
    data = _json()
    df, inferred = _frame(data)
    field = data.get("field")
    if not isinstance(field, str):
        raise AppError("Field tidak valid.", 400)
    search = str(data.get("search") or "")[:100]
    limit = min(1000, max(1, int(data.get("limit", 300))))
    return ok(**ps.field_values(df, inferred, _types(data), field, search, limit))


@bp.post("/api/pivot/calculate")
def calculate():
    data = _json()
    result, _, _ = ps.run_pivot(_sid(), _updir(), data, current_app.config["PIVOT_MAX_CELLS"])
    page = data.get("page") or {}
    try:
        offset = max(0, int(page.get("offset", 0)))
        limit = min(2000, max(1, int(page.get("limit", current_app.config["PIVOT_PAGE_SIZE"]))))
    except (TypeError, ValueError):
        raise AppError("Parameter halaman tidak valid.", 400)
    body = result["body"]
    return ok(
        result={
            "header": result["header"],
            "body": body[offset : offset + limit],
            "meta": result["meta"],
            "total_rows": len(body),
            "offset": offset,
        }
    )


@bp.post("/api/pivot/export")
def export():
    data = _json()
    result, df, cfg = ps.run_pivot(_sid(), _updir(), data, None)
    try:
        decimals = int(data.get("decimals", 2))
    except (TypeError, ValueError):
        decimals = 2
    wb = xs.write_pivot_workbook(result, df, ps.describe_filters(cfg), decimals)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=XLSX_MIME, as_attachment=True, download_name="pivot_table_export.xlsx")
