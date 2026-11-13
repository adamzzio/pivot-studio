import time
import uuid
from pathlib import Path
from urllib.parse import urlparse

from flask import Flask, redirect, request, session, url_for
from werkzeug.exceptions import HTTPException

from config import Config

from .utils import AppError, fail


def create_app(config_object=Config) -> Flask:
    app = Flask(__name__, instance_path=str(config_object.INSTANCE_DIR))
    app.config.from_object(config_object)

    upload_dir = Path(app.config["UPLOAD_DIR"])
    upload_dir.mkdir(parents=True, exist_ok=True)

    from .services import excel_service

    ttl = app.config["FILE_TTL_HOURS"] * 3600
    excel_service.cleanup_stale(upload_dir, ttl)
    last_cleanup = {"t": time.time()}

    @app.before_request
    def _guard():
        # Tolak request lintas-origin untuk metode yang mengubah state.
        if request.method in ("POST", "PUT", "PATCH", "DELETE"):
            origin = request.headers.get("Origin")
            if origin and urlparse(origin).netloc != request.host:
                return fail("Permintaan ditolak karena origin tidak sesuai.", 403)
        if "sid" not in session:
            session["sid"] = uuid.uuid4().hex
        if time.time() - last_cleanup["t"] > 1800:
            last_cleanup["t"] = time.time()
            excel_service.cleanup_stale(upload_dir, ttl)

    @app.after_request
    def _headers(resp):
        resp.headers.setdefault("X-Content-Type-Options", "nosniff")
        resp.headers.setdefault("Referrer-Policy", "same-origin")
        return resp

    from .routes import converter, insights, pivot

    app.register_blueprint(converter.bp)
    app.register_blueprint(pivot.bp)
    app.register_blueprint(insights.bp)

    @app.get("/")
    def index():
        return redirect(url_for("converter.page"))

    @app.errorhandler(AppError)
    def _app_error(exc: AppError):
        return fail(exc.message, exc.status)

    @app.errorhandler(Exception)
    def _unhandled(exc: Exception):
        is_api = request.path.startswith("/api/")
        if isinstance(exc, HTTPException):
            if is_api:
                msg = "File terlalu besar." if exc.code == 413 else (exc.description or exc.name)
                return fail(msg, exc.code or 500)
            return exc
        app.logger.exception("Kesalahan tak terduga")
        if is_api:
            return fail("Terjadi kesalahan tak terduga di server.", 500)
        return "Terjadi kesalahan tak terduga di server.", 500

    return app
