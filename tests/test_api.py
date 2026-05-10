import pytest
from doc_intelligence.web.app import create_app, _doc_cache
from doc_intelligence.engine import ParsedDocument, CellData


@pytest.fixture
def client():
    _doc_cache.clear()
    app, socketio = create_app(testing=True, db_path=":memory:")
    with app.test_client() as c:
        yield c
    _doc_cache.clear()


def test_get_documents_empty(client):
    resp = client.get("/api/documents")
    assert resp.status_code == 200
    assert resp.get_json() == []


def test_get_documents_with_cache(client):
    _doc_cache["test123"] = {
        "info": {"app": "Excel.Application", "name": "a.xlsx", "path": "C:/a.xlsx"},
        "parsed": None,
        "fingerprint": {"labels": ["품명"], "label_positions": {}},
        "match": {"template": None, "score": 0.1, "auto": False},
        "snapshot_b64": None,
        "confirmed": False,
    }
    resp = client.get("/api/documents")
    data = resp.get_json()
    assert len(data) == 1
    assert data[0]["id"] == "test123"
    assert data[0]["status"] == "new"


def test_get_parsed_not_found(client):
    resp = client.get("/api/documents/nonexistent/parsed")
    assert resp.status_code == 404


def test_get_parsed_success(client):
    cells = [CellData(address="Sheet1!R1C1", value="품명", data_type="text", neighbors={})]
    parsed = ParsedDocument(
        file_path="C:/a.xlsx", file_type="excel", raw_text="품명",
        structure={"sheets": ["Sheet1"]}, cells=cells, metadata={}
    )
    _doc_cache["test123"] = {
        "info": {"app": "Excel.Application", "name": "a.xlsx", "path": "C:/a.xlsx"},
        "parsed": parsed,
        "fingerprint": {"labels": ["품명"], "label_positions": {"품명": "Sheet1!R1C1"}},
        "match": {"template": None, "score": 0.1, "auto": False},
        "snapshot_b64": None,
        "confirmed": False,
    }
    resp = client.get("/api/documents/test123/parsed")
    assert resp.status_code == 200
    data = resp.get_json()
    assert len(data["cells"]) == 1
    assert data["cells"][0]["value"] == "품명"


def test_learn_template(client):
    cells = [CellData(address="R1C1", value="품명", data_type="text", neighbors={})]
    parsed = ParsedDocument(
        file_path="C:/a.xlsx", file_type="excel", raw_text="품명",
        structure={}, cells=cells, metadata={}
    )
    _doc_cache["test123"] = {
        "info": {"app": "Excel.Application", "name": "a.xlsx", "path": "C:/a.xlsx"},
        "parsed": parsed,
        "fingerprint": {"labels": ["품명"], "label_positions": {"품명": "R1C1"}},
        "match": {"template": None, "score": 0.1, "auto": False},
        "snapshot_b64": None,
        "confirmed": False,
    }
    resp = client.post("/api/templates/learn", json={
        "doc_id": "test123",
        "template_name": "테스트 양식"
    })
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["template_id"] is not None
    assert _doc_cache["test123"]["match"]["auto"] is True


def test_get_templates(client):
    resp = client.get("/api/templates")
    assert resp.status_code == 200
    assert isinstance(resp.get_json(), list)
