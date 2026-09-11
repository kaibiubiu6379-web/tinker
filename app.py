#!/usr/bin/env python3
import hmac
import os
import re
import secrets
from datetime import datetime, timedelta
from functools import wraps

from flask import Flask, jsonify, redirect, render_template, request, send_file, session, url_for

from domain_to_excel import create_excel_bytes, parse_text


MAX_UPLOAD_BYTES = 2 * 1024 * 1024
ALLOWED_EXTENSIONS = {".txt", ".log"}
EXCEL_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def create_app(test_config=None):
    app = Flask(__name__)
    app.config.from_mapping(
        SECRET_KEY=os.environ.get("TINKER_SECRET_KEY") or secrets.token_hex(32),
        # Leave room for multipart headers; the file itself is checked separately.
        MAX_CONTENT_LENGTH=MAX_UPLOAD_BYTES + 64 * 1024,
        PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=os.environ.get("TINKER_SECURE_COOKIE", "0") == "1",
        APP_USERNAME=os.environ.get("TINKER_USERNAME", "admin"),
        APP_PASSWORD=os.environ.get("TINKER_PASSWORD"),
    )
    if test_config:
        app.config.update(test_config)
    if not app.config["APP_PASSWORD"]:
        raise RuntimeError("必须通过 TINKER_PASSWORD 环境变量设置登录密码")

    def is_authenticated():
        return session.get("authenticated") is True

    def login_required(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if not is_authenticated():
                if request.path.startswith("/api/"):
                    return jsonify(error="登录状态已失效，请重新登录"), 401
                return redirect(url_for("login_page"))
            return view(*args, **kwargs)

        return wrapped

    def csrf_required(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            expected = session.get("csrf_token", "")
            provided = request.headers.get("X-CSRF-Token", "")
            if not expected or not hmac.compare_digest(expected, provided):
                return jsonify(error="请求校验失败，请刷新页面后重试"), 403
            return view(*args, **kwargs)

        return wrapped

    @app.get("/login")
    def login_page():
        if is_authenticated():
            return redirect(url_for("index"))
        return render_template("login.html")

    @app.get("/")
    @login_required
    def index():
        csrf_token = session.get("csrf_token")
        if not csrf_token:
            csrf_token = secrets.token_urlsafe(32)
            session["csrf_token"] = csrf_token
        return render_template("index.html", csrf_token=csrf_token)

    @app.post("/api/login")
    def login():
        payload = request.get_json(silent=True) or request.form
        username = str(payload.get("username", ""))
        password = str(payload.get("password", ""))
        username_ok = hmac.compare_digest(username, app.config["APP_USERNAME"])
        password_ok = hmac.compare_digest(password, app.config["APP_PASSWORD"])
        if not (username_ok and password_ok):
            return jsonify(error="用户名或密码错误"), 401

        session.clear()
        session.permanent = True
        session["authenticated"] = True
        session["csrf_token"] = secrets.token_urlsafe(32)
        return jsonify(ok=True)

    @app.post("/api/logout")
    @login_required
    @csrf_required
    def logout():
        session.clear()
        return jsonify(ok=True)

    @app.post("/api/convert")
    @login_required
    @csrf_required
    def convert():
        title = request.form.get("title", "域名信息").strip() or "域名信息"
        if len(title) > 80:
            return jsonify(error="标题不能超过 80 个字符"), 400

        uploaded_file = request.files.get("file")
        pasted_text = request.form.get("text", "")

        if uploaded_file and uploaded_file.filename:
            extension = os.path.splitext(uploaded_file.filename)[1].lower()
            if extension not in ALLOWED_EXTENSIONS:
                return jsonify(error="目前仅支持上传 TXT 或 LOG 文本文件"), 400
            raw = uploaded_file.read(MAX_UPLOAD_BYTES + 1)
            if len(raw) > MAX_UPLOAD_BYTES:
                return jsonify(error="文件不能超过 2 MB"), 413
            try:
                text = decode_uploaded_text(raw)
            except UnicodeDecodeError:
                return jsonify(error="无法识别文件编码，请使用 UTF-8 或 GB18030 文本"), 400
        elif pasted_text.strip():
            text = pasted_text
        else:
            return jsonify(error="请选择文本文件，或粘贴需要整理的内容"), 400

        records = parse_text(text)
        if not records:
            return jsonify(error="没有识别到域名，请检查文本格式"), 422

        output, categories, _ = create_excel_bytes(records, title)
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        download_name = f"{safe_filename(title)}-{timestamp}.xlsx"
        response = send_file(
            output,
            mimetype=EXCEL_MIMETYPE,
            as_attachment=True,
            download_name=download_name,
        )
        response.headers["X-Record-Count"] = str(len(records))
        response.headers["X-Category-Count"] = str(len(categories))
        return response

    @app.errorhandler(413)
    def file_too_large(_error):
        if request.path.startswith("/api/"):
            return jsonify(error="上传内容不能超过 2 MB"), 413
        return "上传内容不能超过 2 MB", 413

    return app


def decode_uploaded_text(raw):
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("text", raw, 0, len(raw), "unsupported encoding")


def safe_filename(value):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", value).strip(" .")
    return cleaned[:80] or "域名信息"


if __name__ == "__main__":
    app = create_app()
    app.run(host="127.0.0.1", port=5000, debug=False)
