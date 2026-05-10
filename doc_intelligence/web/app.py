"""Flask + SocketIO 서버 + COM 폴링 스레드"""
import hashlib
import logging
import os
import threading
import time
from dataclasses import asdict

from flask import Flask
from flask_socketio import SocketIO

from doc_intelligence.com_worker import ComWorker
from doc_intelligence.parsers import ExcelParser, WordParser, PowerPointParser, PdfParser
from doc_intelligence.fingerprint import Fingerprinter
from doc_intelligence.engine import Engine
from doc_intelligence.web.snapshot import capture_window_snapshot

logger = logging.getLogger(__name__)

_APP_TO_PARSER = {
    "Excel.Application": ExcelParser,
    "Word.Application": WordParser,
    "PowerPoint.Application": PowerPointParser,
    "AcroExch.App": PdfParser,
}

_doc_cache: dict[str, dict] = {}
_polling_running = False


def _make_doc_id(file_path: str) -> str:
    return hashlib.md5(file_path.encode("utf-8")).hexdigest()


def _build_doc_summary(doc_id: str, entry: dict, template_names: dict) -> dict:
    match = entry["match"]
    score = match.get("score", 0.0)
    template_id = match.get("template")
    auto = match.get("auto", False)
    confirmed = entry.get("confirmed", False)

    if auto or confirmed:
        status = "matched"
    elif template_id is not None and score >= 0.60:
        status = "candidate"
    else:
        status = "new"

    info = entry["info"]
    return {
        "id": doc_id,
        "app": info.get("app", ""),
        "name": info.get("name", ""),
        "path": info.get("path", ""),
        "status": status,
        "score": round(score * 100, 1) if score else 0,
        "template_id": template_id,
        "template_name": template_names.get(template_id) if template_id else None,
        "labels": entry.get("fingerprint", {}).get("labels", []),
        "has_preview": entry.get("snapshot_b64") is not None,
    }


def _polling_loop(com_worker, engine, fingerprinter, socketio, interval=3):
    global _polling_running
    _polling_running = True
    print("[polling] thread started", flush=True)

    import pythoncom
    try:
        pythoncom.CoInitialize()
        print("[polling] CoInitialize OK", flush=True)
    except Exception as e:
        print(f"[polling] CoInitialize FAIL: {e}", flush=True)
        return

    try:
        while _polling_running:
            try:
                docs = com_worker.detect_open_documents()
                print(f"[polling] detected: {len(docs)}", flush=True)
                current_ids = set()
                for doc_info in docs:
                    file_path = doc_info.get("path", "")
                    if not file_path:
                        continue
                    doc_id = _make_doc_id(file_path)
                    current_ids.add(doc_id)
                    if doc_id in _doc_cache and _doc_cache[doc_id].get("confirmed"):
                        continue
                    if doc_id in _doc_cache and _doc_cache[doc_id].get("parsed"):
                        continue
                    app_type = doc_info.get("app", "")
                    parser_cls = _APP_TO_PARSER.get(app_type)
                    if parser_cls is None:
                        continue
                    try:
                        parser = parser_cls()
                        com_app = doc_info.get("app_obj")
                        parsed = parser.parse_from_com(com_app)
                        fp_result = fingerprinter.generate(parsed)
                        match_result = fingerprinter.match(parsed)
                        snapshot = capture_window_snapshot(doc_info.get("name", ""))
                        _doc_cache[doc_id] = {
                            "info": {k: v for k, v in doc_info.items() if k != "app_obj"},
                            "parsed": parsed,
                            "fingerprint": fp_result,
                            "match": match_result,
                            "snapshot_b64": snapshot,
                            "confirmed": False,
                        }
                        socketio.emit("parse_complete", {"doc_id": doc_id, "status": "ok"})
                        socketio.emit("documents_updated", _get_all_summaries(engine))
                        print(f"[polling] processed: {doc_info.get('name')}", flush=True)
                    except Exception as e:
                        print(f"[polling] parse fail ({file_path}): {e}", flush=True)
                        import traceback
                        traceback.print_exc()
                closed = set(_doc_cache.keys()) - current_ids
                if closed:
                    for cid in closed:
                        del _doc_cache[cid]
                    socketio.emit("documents_updated", _get_all_summaries(engine))
            except Exception as e:
                print(f"[polling] loop error: {e}", flush=True)
            time.sleep(interval)
    finally:
        pythoncom.CoUninitialize()


def _get_all_summaries(engine):
    templates = engine.storage.get_all_templates()
    template_names = {t["id"]: t["name"] for t in templates}
    return [_build_doc_summary(doc_id, entry, template_names) for doc_id, entry in _doc_cache.items()]


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
    api_bp = create_api_blueprint(engine, fingerprinter, _doc_cache)
    app.register_blueprint(api_bp, url_prefix="/api")
    if not testing:
        polling_thread = threading.Thread(
            target=_polling_loop,
            args=(com_worker, engine, fingerprinter, socketio),
            daemon=True,
        )
        polling_thread.start()
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
