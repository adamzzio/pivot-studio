from flask import jsonify


class AppError(Exception):
    """Error yang aman ditampilkan ke pengguna."""

    def __init__(self, message: str, status: int = 400):
        super().__init__(message)
        self.message = message
        self.status = status


def ok(**data):
    return jsonify({"ok": True, **data})


def fail(message: str, status: int = 400):
    return jsonify({"ok": False, "error": message}), status
