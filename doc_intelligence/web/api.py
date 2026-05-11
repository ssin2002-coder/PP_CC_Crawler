"""REST API 엔드포인트"""
import base64
import hashlib
import io
import os
from dataclasses import asdict
from flask import Blueprint, Response, jsonify, request
from doc_intelligence.engine import Engine
from doc_intelligence.fingerprint import Fingerprinter
from doc_intelligence.parsers import ImageParser


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


def create_api_blueprint(engine: Engine, fingerprinter: Fingerprinter, doc_cache: dict):
    api = Blueprint("api", __name__)

    def _template_names():
        templates = engine.storage.get_all_templates()
        return {t["id"]: t["name"] for t in templates}

    def _build_summary(doc_id, entry):
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
        names = _template_names()
        info = entry["info"]
        return {
            "id": doc_id, "app": info.get("app", ""), "name": info.get("name", ""),
            "path": info.get("path", ""), "status": status,
            "score": round(score * 100, 1) if score else 0,
            "template_id": template_id,
            "template_name": names.get(template_id) if template_id else None,
            "labels": entry.get("fingerprint", {}).get("labels", []),
            "has_preview": entry.get("snapshot_b64") is not None,
        }

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
        entry["match"]["auto"] = True
        return jsonify({"status": "confirmed"})

    @api.route("/templates", methods=["GET"])
    def get_templates():
        templates = engine.storage.get_all_templates()
        return jsonify(templates)

    @api.route("/documents/upload-image", methods=["POST"])
    def upload_image():
        """이미지 파일 경로를 받아 OCR 파싱 후 doc_cache에 등록한다."""
        data = request.get_json()
        file_path = data.get("file_path") if data else None
        if not file_path:
            return jsonify({"error": "valid file_path required"}), 400
        file_path = os.path.abspath(file_path)
        if not os.path.isfile(file_path):
            return jsonify({"error": "file not found"}), 400

        doc_id = hashlib.md5(file_path.encode("utf-8")).hexdigest()
        if doc_id in doc_cache and doc_cache[doc_id].get("parsed"):
            return jsonify({"doc_id": doc_id, "status": "already_parsed"})

        parser = ImageParser()
        parsed = parser.parse_from_com(file_path)
        fp_result = fingerprinter.generate(parsed)
        match_result = fingerprinter.match(parsed)

        doc_cache[doc_id] = {
            "info": {"app": "Image", "name": os.path.basename(file_path), "path": file_path},
            "parsed": parsed,
            "fingerprint": fp_result,
            "match": match_result,
            "snapshot_b64": None,
            "confirmed": False,
            "source": "api",
        }
        return jsonify({"doc_id": doc_id, "status": "parsed"})

    @api.route("/documents/upload-pdf", methods=["POST"])
    def upload_pdf():
        """PDF 파일 경로를 받아 pypdf로 실제 파싱 후 doc_cache에 등록한다."""
        data = request.get_json()
        file_path = data.get("file_path") if data else None
        if not file_path:
            return jsonify({"error": "valid file_path required"}), 400
        file_path = os.path.abspath(file_path)
        if not os.path.isfile(file_path):
            return jsonify({"error": "file not found"}), 400

        doc_id = hashlib.md5(file_path.encode("utf-8")).hexdigest()
        if doc_id in doc_cache and doc_cache[doc_id].get("parsed"):
            return jsonify({"doc_id": doc_id, "status": "already_parsed"})

        from doc_intelligence.parsers import PdfParser
        parsed = PdfParser.parse_from_file(file_path)

        if parsed.metadata.get("fallback"):
            doc_cache[doc_id] = {
                "info": {"app": "AcroExch.App", "name": os.path.basename(file_path), "path": file_path},
                "parsed": parsed,
                "fingerprint": {"labels": []},
                "match": {"template": None, "score": 0.0, "auto": False},
                "snapshot_b64": None,
                "confirmed": False,
                "source": "api",
            }
            return jsonify({"doc_id": doc_id, "status": "parse_failed", "reason": parsed.metadata.get("reason", "")}), 200

        fp_result = fingerprinter.generate(parsed)
        match_result = fingerprinter.match(parsed)

        pdf_preview = _render_pdf_preview(file_path)
        doc_cache[doc_id] = {
            "info": {"app": "AcroExch.App", "name": os.path.basename(file_path), "path": file_path},
            "parsed": parsed,
            "fingerprint": fp_result,
            "match": match_result,
            "snapshot_b64": pdf_preview,
            "confirmed": False,
            "source": "api",
        }
        return jsonify({"doc_id": doc_id, "status": "parsed"})

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
