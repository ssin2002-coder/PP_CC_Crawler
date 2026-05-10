"""REST API 엔드포인트"""
import base64
from dataclasses import asdict
from flask import Blueprint, Response, jsonify, request
from doc_intelligence.engine import Engine
from doc_intelligence.fingerprint import Fingerprinter


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

    return api
