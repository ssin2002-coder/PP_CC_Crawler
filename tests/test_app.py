import pytest
from unittest.mock import patch, MagicMock
from doc_intelligence.web.app import create_app, _build_doc_summary


def test_create_app_returns_flask_app():
    app, socketio = create_app(testing=True, db_path=":memory:")
    assert app is not None
    assert app.testing is True


def test_build_doc_summary_matched():
    cache_entry = {
        "info": {"app": "Excel.Application", "name": "test.xlsx", "path": "C:/test.xlsx"},
        "match": {"template": 1, "score": 0.92, "auto": True},
        "fingerprint": {"labels": ["품명", "단가"], "label_positions": {}},
        "snapshot_b64": "abc123",
        "confirmed": False,
    }
    result = _build_doc_summary("abc123hash", cache_entry, template_names={1: "정산서 양식"})
    assert result["id"] == "abc123hash"
    assert result["status"] == "matched"
    assert result["template_name"] == "정산서 양식"
    assert result["score"] == 92.0


def test_build_doc_summary_candidate():
    cache_entry = {
        "info": {"app": "Word.Application", "name": "doc.docx", "path": "C:/doc.docx"},
        "match": {"template": 2, "score": 0.73, "auto": False},
        "fingerprint": {"labels": ["항목"], "label_positions": {}},
        "snapshot_b64": None,
        "confirmed": False,
    }
    result = _build_doc_summary("def456hash", cache_entry, template_names={2: "일보 양식"})
    assert result["status"] == "candidate"


def test_build_doc_summary_new():
    cache_entry = {
        "info": {"app": "Excel.Application", "name": "new.xlsx", "path": "C:/new.xlsx"},
        "match": {"template": None, "score": 0.3, "auto": False},
        "fingerprint": {"labels": [], "label_positions": {}},
        "snapshot_b64": None,
        "confirmed": False,
    }
    result = _build_doc_summary("ghi789hash", cache_entry, template_names={})
    assert result["status"] == "new"
    assert result["template_name"] is None
