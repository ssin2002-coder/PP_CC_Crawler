"""Flask + SocketIO 서버 — 명시적 스캔/파싱 + 클라이언트 종료 시 자동 셧다운"""
import hashlib
import logging
import os
import threading

from flask import Flask, request as flask_request
from flask_socketio import SocketIO

from doc_intelligence.com_worker import ComWorker
from doc_intelligence.fingerprint import Fingerprinter
from doc_intelligence.engine import Engine

logger = logging.getLogger(__name__)

_doc_cache: dict[str, dict] = {}

_connected_clients: set[str] = set()
_shutdown_timer: threading.Timer | None = None
_shutdown_lock = threading.Lock()
_SHUTDOWN_GRACE = 5.0


def _make_doc_id(file_path: str) -> str:
    return hashlib.md5(file_path.encode("utf-8")).hexdigest()


def _exit():
    with _shutdown_lock:
        if _connected_clients:
            return
    os._exit(0)


def _arm_shutdown():
    global _shutdown_timer
    with _shutdown_lock:
        if _connected_clients:
            return
        if _shutdown_timer is not None and _shutdown_timer.is_alive():
            return
        _shutdown_timer = threading.Timer(_SHUTDOWN_GRACE, _exit)
        _shutdown_timer.daemon = True
        _shutdown_timer.start()


def _cancel_shutdown():
    global _shutdown_timer
    with _shutdown_lock:
        if _shutdown_timer is not None:
            _shutdown_timer.cancel()
            _shutdown_timer = None


def create_app(testing=False, db_path="doc_intelligence.db"):
    app = Flask(__name__, static_folder=None)
    app.testing = testing
    socketio = SocketIO(app, cors_allowed_origins="*", async_mode="threading")
    engine = Engine(db_path=db_path)
    fingerprinter = Fingerprinter(storage=engine.storage)
    fingerprinter.initialize(engine)
    com_worker = ComWorker()
    app.config["engine"] = engine
    app.config["fingerprinter"] = fingerprinter
    app.config["com_worker"] = com_worker
    from doc_intelligence.web.api import create_api_blueprint
    api_bp = create_api_blueprint(engine, fingerprinter, _doc_cache, socketio)
    app.register_blueprint(api_bp, url_prefix="/api")

    if not testing:
        @socketio.on("connect")
        def _on_connect():
            sid = flask_request.sid
            with _shutdown_lock:
                _connected_clients.add(sid)
            _cancel_shutdown()

        @socketio.on("disconnect")
        def _on_disconnect():
            sid = flask_request.sid
            with _shutdown_lock:
                _connected_clients.discard(sid)
            _arm_shutdown()

    static_dir = os.path.join(os.path.dirname(__file__), "static")
    if os.path.exists(static_dir):
        from flask import send_from_directory

        @app.route("/")
        def index():
            return send_from_directory(static_dir, "index.html")

        @app.route("/<path:path>")
        def static_files(path):
            file_path = os.path.join(static_dir, path)
            if os.path.exists(file_path):
                return send_from_directory(static_dir, path)
            return send_from_directory(static_dir, "index.html")

    return app, socketio
