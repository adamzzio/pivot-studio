from flask import Blueprint, current_app, request, session

from app.services import groq_service as gs
from app.services import pivot_service as ps
from app.utils import AppError, ok

bp = Blueprint("insights", __name__)


@bp.post("/api/ai/insights")
def insights():
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        raise AppError("Body permintaan harus berupa JSON.", 400)
    cfg_app = current_app.config
    if not cfg_app["GROQ_API_KEY"]:
        raise AppError("GROQ_API_KEY belum diatur. Isi pada file .env lalu jalankan ulang aplikasi.", 503)

    opts = data.get("options") or {}
    language = opts.get("language") if opts.get("language") in gs.LANGS else "id"
    length = opts.get("length") if opts.get("length") in gs.LENGTHS else "ringkas"

    # Hasil dihitung ulang di server agar angka yang dikirim ke AI selalu berasal dari data asli.
    result, _, cfg = ps.run_pivot(session["sid"], cfg_app["UPLOAD_DIR"], data, cfg_app["PIVOT_MAX_CELLS"])
    if not result["body"]:
        raise AppError("Hasil pivot kosong. Tidak ada yang dapat dianalisis.", 422)
    text, truncated = gs.build_payload(result, cfg, cfg_app["AI_MAX_ROWS"], cfg_app["AI_MAX_COLS"])
    out = gs.generate_insight(
        text,
        api_key=cfg_app["GROQ_API_KEY"],
        model=cfg_app["GROQ_MODEL"],
        language=language,
        length=length,
        timeout=cfg_app["GROQ_TIMEOUT"],
    )
    return ok(truncated=truncated, model=cfg_app["GROQ_MODEL"], **out)
