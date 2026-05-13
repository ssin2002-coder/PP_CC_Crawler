"""REST API 엔드포인트 — 명시적 스캔/파싱 파이프라인"""
import base64
import hashlib
import os
from dataclasses import asdict
from flask import Blueprint, Response, current_app, jsonify, request
import yaml
from doc_intelligence.engine import Engine
from doc_intelligence.fingerprint import Fingerprinter
from doc_intelligence.parsers import (
    ExcelParser, WordParser, PowerPointParser, PdfParser, ImageParser,
)
from doc_intelligence.web.snapshot import capture_window_snapshot


_COM_PARSERS = {
    "Excel.Application": ExcelParser,
    "Word.Application": WordParser,
    "PowerPoint.Application": PowerPointParser,
    "AcroExch.App": PdfParser,
}


def _render_pdf_preview(file_path: str) -> str | None:
    """PyMuPDF로 PDF 첫 페이지를 PNG base64로 렌더링한다."""
    try:
        import fitz
        doc = fitz.open(file_path)
        page = doc[0]
        pix = page.get_pixmap(dpi=120)
        img_bytes = pix.tobytes("png")
        doc.close()
        return base64.b64encode(img_bytes).decode("utf-8")
    except Exception:
        return None


def _make_doc_id(file_path: str) -> str:
    return hashlib.md5(file_path.encode("utf-8")).hexdigest()


def _load_watch_dirs():
    config_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.yaml"
    )
    image_dirs: list = []
    pdf_dirs: list = []
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
        image_dirs = cfg.get("image", {}).get("watch_dirs", []) or []
        pdf_dirs = cfg.get("pdf", {}).get("watch_dirs", []) or []
    except Exception as exc:
        print(f"[scan] config load failed: {exc}")
    return image_dirs, pdf_dirs


def create_api_blueprint(engine: Engine, fingerprinter: Fingerprinter, doc_cache: dict, socketio=None):
    api = Blueprint("api", __name__)

    def _template_names():
        templates = engine.storage.get_all_templates()
        return {t["id"]: t["name"] for t in templates}

    def _build_summary(doc_id, entry):
        info = entry["info"]
        parsed_state = entry.get("parsed_state", "parsed")
        match = entry.get("match") or {}
        score = match.get("score", 0.0)
        template_id = match.get("template")
        auto = match.get("auto", False)
        confirmed = entry.get("confirmed", False)
        if parsed_state == "parsed":
            if auto or confirmed:
                status = "matched"
            elif template_id is not None and score >= 0.60:
                status = "candidate"
            else:
                status = "new"
        else:
            status = "new"
        names = _template_names()
        return {
            "id": doc_id,
            "app": info.get("app", ""),
            "name": info.get("name", ""),
            "path": info.get("path", ""),
            "status": status,
            "parsed_state": parsed_state,
            "score": round(score * 100, 1) if score else 0,
            "template_id": template_id,
            "template_name": names.get(template_id) if template_id else None,
            "labels": (entry.get("fingerprint") or {}).get("labels", []),
            "has_preview": entry.get("snapshot_b64") is not None,
            "error": entry.get("error"),
        }

    def _emit_summaries():
        if socketio is None:
            return
        summaries = [_build_summary(did, e) for did, e in doc_cache.items()]
        socketio.emit("documents_updated", summaries)

    @api.route("/documents", methods=["GET"])
    def get_documents():
        summaries = [_build_summary(did, e) for did, e in doc_cache.items()]
        return jsonify(summaries)

    @api.route("/documents/<doc_id>/preview", methods=["GET"])
    def get_preview(doc_id):
        entry = doc_cache.get(doc_id)
        if entry is None:
            return jsonify({"error": "not found"}), 404
        b64 = entry.get("snapshot_b64")
        if b64 is None:
            return jsonify({"error": "no preview"}), 404
        img_bytes = base64.b64decode(b64)
        return Response(img_bytes, mimetype="image/png")

    @api.route("/documents/<doc_id>/parsed", methods=["GET"])
    def get_parsed(doc_id):
        entry = doc_cache.get(doc_id)
        if entry is None:
            return jsonify({"error": "not found"}), 404
        if entry.get("parsed_state") != "parsed":
            return jsonify({"error": "not parsed", "parsed_state": entry.get("parsed_state")}), 404
        parsed = entry.get("parsed")
        if parsed is None:
            return jsonify({"error": "not parsed"}), 404
        cells_data = [asdict(c) for c in parsed.cells]
        return jsonify({
            "file_path": parsed.file_path, "file_type": parsed.file_type,
            "structure": parsed.structure, "cells": cells_data, "metadata": parsed.metadata,
        })

    @api.route("/templates/learn", methods=["POST"])
    def learn_template():
        data = request.get_json()
        doc_id = data.get("doc_id")
        template_name = data.get("template_name")
        if not doc_id or not template_name:
            return jsonify({"error": "doc_id and template_name required"}), 400
        entry = doc_cache.get(doc_id)
        if entry is None or entry.get("parsed") is None:
            return jsonify({"error": "document not found or not parsed"}), 404
        parsed = entry["parsed"]
        template_id = fingerprinter.learn(parsed, template_name)
        entry["match"] = {"template": template_id, "score": 1.0, "auto": True}
        return jsonify({"template_id": template_id, "status": "learned"})

    @api.route("/templates/confirm", methods=["POST"])
    def confirm_template():
        data = request.get_json()
        doc_id = data.get("doc_id")
        template_id = data.get("template_id")
        if not doc_id or template_id is None:
            return jsonify({"error": "doc_id and template_id required"}), 400
        entry = doc_cache.get(doc_id)
        if entry is None:
            return jsonify({"error": "document not found"}), 404
        engine.storage.increment_match_count(template_id)
        entry["confirmed"] = True
        if entry.get("match") is None:
            entry["match"] = {"template": template_id, "score": 1.0, "auto": True}
        else:
            entry["match"]["auto"] = True
        return jsonify({"status": "confirmed"})

    @api.route("/templates", methods=["GET"])
    def get_templates():
        templates = engine.storage.get_all_templates()
        return jsonify(templates)

    @api.route("/documents/scan", methods=["POST"])
    def scan_documents():
        """열려있는 COM 문서와 감시 폴더의 이미지/PDF를 탐지하기만 한다 (파싱 없음)."""
        image_dirs, pdf_dirs = _load_watch_dirs()
        com_worker = current_app.config.get("com_worker")
        if com_worker is None:
            return jsonify({"error": "com_worker not configured"}), 500

        com_docs: list = []
        try:
            from doc_intelligence.com_worker import _COM_AVAILABLE
        except Exception:
            _COM_AVAILABLE = False
        if _COM_AVAILABLE:
            try:
                import pythoncom
                pythoncom.CoInitialize()
                try:
                    com_docs = com_worker.detect_open_documents()
                finally:
                    pythoncom.CoUninitialize()
            except Exception as exc:
                print(f"[scan] COM detect failed: {exc}")

        image_docs = com_worker.detect_image_files(image_dirs) if image_dirs else []
        pdf_docs = com_worker.detect_pdf_files(pdf_dirs) if pdf_dirs else []

        found_ids: set = set()

        for doc in com_docs:
            file_path = doc.get("path", "")
            if not file_path:
                continue
            doc_id = _make_doc_id(file_path)
            found_ids.add(doc_id)
            if doc_id in doc_cache:
                continue
            doc_cache[doc_id] = {
                "info": {"app": doc.get("app", ""), "name": doc.get("name", ""), "path": file_path},
                "parsed_state": "discovered",
                "source_type": "com",
                "parsed": None,
                "fingerprint": None,
                "match": None,
                "snapshot_b64": None,
                "confirmed": False,
                "source": "scan",
                "error": None,
            }

        for doc in image_docs:
            file_path = doc.get("path", "")
            if not file_path:
                continue
            doc_id = _make_doc_id(file_path)
            found_ids.add(doc_id)
            if doc_id in doc_cache:
                continue
            doc_cache[doc_id] = {
                "info": {"app": "Image", "name": doc.get("name", ""), "path": file_path},
                "parsed_state": "discovered",
                "source_type": "image_file",
                "parsed": None,
                "fingerprint": None,
                "match": None,
                "snapshot_b64": None,
                "confirmed": False,
                "source": "scan",
                "error": None,
            }

        for doc in pdf_docs:
            file_path = doc.get("path", "")
            if not file_path:
                continue
            doc_id = _make_doc_id(file_path)
            found_ids.add(doc_id)
            if doc_id in doc_cache:
                continue
            doc_cache[doc_id] = {
                "info": {"app": "AcroExch.App", "name": doc.get("name", ""), "path": file_path},
                "parsed_state": "discovered",
                "source_type": "pdf_file",
                "parsed": None,
                "fingerprint": None,
                "match": None,
                "snapshot_b64": None,
                "confirmed": False,
                "source": "scan",
                "error": None,
            }

        stale = [
            did for did, entry in doc_cache.items()
            if entry.get("parsed_state") == "discovered" and did not in found_ids
        ]
        for did in stale:
            del doc_cache[did]

        summaries = [_build_summary(did, e) for did, e in doc_cache.items()]
        if socketio is not None:
            socketio.emit("documents_updated", summaries)
        return jsonify({"detected": len(found_ids), "documents": summaries}), 200

    @api.route("/documents/<doc_id>/parse", methods=["POST"])
    def parse_document(doc_id):
        """단일 문서를 요청 시점에 파싱한다."""
        entry = doc_cache.get(doc_id)
        if entry is None:
            return jsonify({"error": "not found"}), 404
        if entry.get("parsed_state") == "parsed":
            return jsonify({"doc_id": doc_id, "parsed_state": "parsed", "status": "already_parsed"}), 200

        entry["parsed_state"] = "parsing"
        entry["error"] = None
        _emit_summaries()

        source_type = entry.get("source_type")
        file_path = entry["info"].get("path", "")
        com_worker = current_app.config.get("com_worker")

        try:
            if source_type == "image_file":
                parsed = ImageParser().parse_from_com(file_path)
                snapshot = None
            elif source_type == "pdf_file":
                parsed = PdfParser.parse_from_file(file_path)
                snapshot = _render_pdf_preview(file_path)
            elif source_type == "com":
                try:
                    from doc_intelligence.com_worker import _COM_AVAILABLE
                except Exception:
                    _COM_AVAILABLE = False
                if not _COM_AVAILABLE:
                    raise RuntimeError("COM not available on this platform")
                import pythoncom
                pythoncom.CoInitialize()
                try:
                    docs = com_worker.detect_open_documents()
                    target = None
                    for d in docs:
                        if d.get("path", "") == file_path:
                            target = d
                            break
                    if target is None:
                        raise RuntimeError(f"COM document no longer open: {file_path}")
                    app_type = target.get("app", "")
                    parser_cls = _COM_PARSERS.get(app_type)
                    if parser_cls is None:
                        raise RuntimeError(f"no parser for app type: {app_type}")
                    parser = parser_cls()
                    com_app = target.get("app_obj")
                    doc_obj = target.get("doc_obj")
                    if app_type == "AcroExch.App":
                        pd_doc = target.get("pd_doc")
                        parsed = parser.parse_from_com(com_app, pd_doc=pd_doc)
                    elif doc_obj is not None:
                        parsed = parser.parse_from_com(com_app, doc_obj=doc_obj)
                    else:
                        parsed = parser.parse_from_com(com_app)
                    snapshot = capture_window_snapshot(target.get("name", ""))
                finally:
                    pythoncom.CoUninitialize()
            else:
                raise RuntimeError(f"unknown source_type: {source_type}")

            if parsed.metadata.get("fallback"):
                entry["parsed"] = parsed
                entry["fingerprint"] = {"labels": []}
                entry["match"] = {"template": None, "score": 0.0, "auto": False}
            else:
                entry["parsed"] = parsed
                entry["fingerprint"] = fingerprinter.generate(parsed)
                entry["match"] = fingerprinter.match(parsed)
            entry["snapshot_b64"] = snapshot
            entry["parsed_state"] = "parsed"
            entry["error"] = None
        except Exception as exc:
            print(f"[parse] failed ({file_path}): {exc}")
            entry["parsed_state"] = "error"
            entry["error"] = str(exc)

        _emit_summaries()
        if socketio is not None:
            socketio.emit("parse_complete", {"doc_id": doc_id, "status": entry["parsed_state"]})
        return jsonify({
            "doc_id": doc_id,
            "parsed_state": entry["parsed_state"],
            "error": entry.get("error"),
        }), 200

    @api.route("/status", methods=["GET"])
    def get_status():
        """환경 상태 반환 — COM, Acrobat, Tesseract, pypdf 가용 여부."""
        from doc_intelligence.com_worker import _COM_AVAILABLE
        from doc_intelligence.parsers import _TESSERACT_AVAILABLE
        acrobat_available = False
        if _COM_AVAILABLE:
            try:
                import win32com.client
                win32com.client.GetActiveObject("AcroExch.App")
                acrobat_available = True
            except Exception:
                pass
        pypdf_available = False
        try:
            import pypdf
            pypdf_available = True
        except ImportError:
            pass
        return jsonify({
            "com_available": _COM_AVAILABLE,
            "acrobat_available": acrobat_available,
            "tesseract_available": _TESSERACT_AVAILABLE,
            "pypdf_available": pypdf_available,
        })

    return api
