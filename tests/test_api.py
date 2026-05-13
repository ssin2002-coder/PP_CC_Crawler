import pytest
from unittest.mock import patch
from doc_intelligence.web.app import create_app, _doc_cache
from doc_intelligence.engine import ParsedDocument, CellData


@pytest.fixture
def client():
    _doc_cache.clear()
    app, socketio = create_app(testing=True, db_path=":memory:")
    with app.test_client() as c:
        yield c
    _doc_cache.clear()


def _new_entry(path="C:/a.xlsx", name="a.xlsx", app="Excel.Application",
               parsed_state="parsed", source_type="com", parsed=None,
               match=None, fingerprint=None):
    return {
        "info": {"app": app, "name": name, "path": path},
        "parsed_state": parsed_state,
        "source_type": source_type,
        "parsed": parsed,
        "fingerprint": fingerprint or {"labels": []},
        "match": match or {"template": None, "score": 0.1, "auto": False},
        "snapshot_b64": None,
        "confirmed": False,
        "source": "scan",
        "error": None,
    }


def test_get_documents_empty(client):
    resp = client.get("/api/documents")
    assert resp.status_code == 200
    assert resp.get_json() == []


def test_get_documents_with_cache(client):
    _doc_cache["test123"] = _new_entry(
        fingerprint={"labels": ["품명"], "label_positions": {}}
    )
    resp = client.get("/api/documents")
    data = resp.get_json()
    assert len(data) == 1
    assert data[0]["id"] == "test123"
    assert data[0]["status"] == "new"
    assert data[0]["parsed_state"] == "parsed"


def test_get_documents_discovered_status_is_new(client):
    _doc_cache["disc1"] = _new_entry(parsed_state="discovered")
    resp = client.get("/api/documents")
    data = resp.get_json()
    assert data[0]["status"] == "new"
    assert data[0]["parsed_state"] == "discovered"


def test_get_parsed_not_found(client):
    resp = client.get("/api/documents/nonexistent/parsed")
    assert resp.status_code == 404


def test_get_parsed_not_parsed_state(client):
    _doc_cache["disc1"] = _new_entry(parsed_state="discovered")
    resp = client.get("/api/documents/disc1/parsed")
    assert resp.status_code == 404


def test_get_parsed_success(client):
    cells = [CellData(address="Sheet1!R1C1", value="품명", data_type="text", neighbors={})]
    parsed = ParsedDocument(
        file_path="C:/a.xlsx", file_type="excel", raw_text="품명",
        structure={"sheets": ["Sheet1"]}, cells=cells, metadata={}
    )
    _doc_cache["test123"] = _new_entry(
        parsed=parsed,
        fingerprint={"labels": ["품명"], "label_positions": {"품명": "Sheet1!R1C1"}},
    )
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
    _doc_cache["test123"] = _new_entry(
        parsed=parsed,
        fingerprint={"labels": ["품명"], "label_positions": {"품명": "R1C1"}},
    )
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


def test_scan_no_documents(client):
    with patch("doc_intelligence.com_worker.ComWorker.detect_open_documents", return_value=[]), \
         patch("doc_intelligence.com_worker.ComWorker.detect_image_files", return_value=[]), \
         patch("doc_intelligence.com_worker.ComWorker.detect_pdf_files", return_value=[]):
        resp = client.post("/api/documents/scan")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["detected"] == 0
    assert body["documents"] == []


def test_scan_inserts_discovered_entries(client):
    img_docs = [{"app": "Image", "name": "y.png", "path": "C:/y.png"}]
    pdf_docs = [{"app": "AcroExch.App", "name": "z.pdf", "path": "C:/z.pdf"}]
    with patch("doc_intelligence.com_worker.ComWorker.detect_open_documents", return_value=[]), \
         patch("doc_intelligence.com_worker.ComWorker.detect_image_files", return_value=img_docs), \
         patch("doc_intelligence.com_worker.ComWorker.detect_pdf_files", return_value=pdf_docs), \
         patch("doc_intelligence.web.api._load_watch_dirs", return_value=(["dummy_img"], ["dummy_pdf"])):
        resp = client.post("/api/documents/scan")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["detected"] == 2
    for entry in _doc_cache.values():
        assert entry["parsed_state"] == "discovered"
        assert entry["parsed"] is None
        assert entry["source"] == "scan"
    source_types = {e["source_type"] for e in _doc_cache.values()}
    assert source_types == {"image_file", "pdf_file"}


def test_scan_does_not_reparse_existing(client):
    _doc_cache["existing"] = _new_entry(path="C:/existing.xlsx", parsed_state="parsed")
    parsed_before = _doc_cache["existing"]["parsed_state"]
    with patch("doc_intelligence.com_worker.ComWorker.detect_open_documents", return_value=[]), \
         patch("doc_intelligence.com_worker.ComWorker.detect_image_files", return_value=[]), \
         patch("doc_intelligence.com_worker.ComWorker.detect_pdf_files", return_value=[]):
        resp = client.post("/api/documents/scan")
    assert resp.status_code == 200
    assert _doc_cache["existing"]["parsed_state"] == parsed_before


def test_scan_evicts_stale_discovered(client):
    _doc_cache["stale"] = _new_entry(parsed_state="discovered", path="C:/gone.xlsx")
    with patch("doc_intelligence.com_worker.ComWorker.detect_open_documents", return_value=[]), \
         patch("doc_intelligence.com_worker.ComWorker.detect_image_files", return_value=[]), \
         patch("doc_intelligence.com_worker.ComWorker.detect_pdf_files", return_value=[]):
        resp = client.post("/api/documents/scan")
    assert resp.status_code == 200
    assert "stale" not in _doc_cache


def test_scan_keeps_parsed_even_if_gone(client):
    _doc_cache["keepme"] = _new_entry(parsed_state="parsed", path="C:/keep.xlsx")
    with patch("doc_intelligence.com_worker.ComWorker.detect_open_documents", return_value=[]), \
         patch("doc_intelligence.com_worker.ComWorker.detect_image_files", return_value=[]), \
         patch("doc_intelligence.com_worker.ComWorker.detect_pdf_files", return_value=[]):
        resp = client.post("/api/documents/scan")
    assert resp.status_code == 200
    assert "keepme" in _doc_cache


def test_parse_unknown_doc(client):
    resp = client.post("/api/documents/nope/parse")
    assert resp.status_code == 404


def test_parse_already_parsed_is_noop(client):
    cells = [CellData(address="R1C1", value="a", data_type="text", neighbors={})]
    parsed = ParsedDocument(
        file_path="C:/a.xlsx", file_type="excel", raw_text="a",
        structure={}, cells=cells, metadata={}
    )
    _doc_cache["d1"] = _new_entry(parsed=parsed, parsed_state="parsed")
    resp = client.post("/api/documents/d1/parse")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["parsed_state"] == "parsed"
    assert body.get("status") == "already_parsed"


def test_parse_image_file(client):
    cells = [CellData(address="R1C1", value="hi", data_type="text", neighbors={})]
    parsed = ParsedDocument(
        file_path="C:/y.png", file_type="image", raw_text="hi",
        structure={}, cells=cells, metadata={}
    )
    _doc_cache["img1"] = _new_entry(
        path="C:/y.png", name="y.png", app="Image",
        parsed_state="discovered", source_type="image_file",
    )
    with patch("doc_intelligence.web.api.ImageParser") as MockImage:
        MockImage.return_value.parse_from_com.return_value = parsed
        resp = client.post("/api/documents/img1/parse")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["parsed_state"] == "parsed"
    assert _doc_cache["img1"]["parsed"] is parsed


def test_parse_pdf_file(client):
    cells = [CellData(address="R1C1", value="x", data_type="text", neighbors={})]
    parsed = ParsedDocument(
        file_path="C:/z.pdf", file_type="pdf", raw_text="x",
        structure={}, cells=cells, metadata={}
    )
    _doc_cache["pdf1"] = _new_entry(
        path="C:/z.pdf", name="z.pdf", app="AcroExch.App",
        parsed_state="discovered", source_type="pdf_file",
    )
    with patch("doc_intelligence.web.api.PdfParser") as MockPdf, \
         patch("doc_intelligence.web.api._render_pdf_preview", return_value=None):
        MockPdf.parse_from_file.return_value = parsed
        resp = client.post("/api/documents/pdf1/parse")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["parsed_state"] == "parsed"
    assert _doc_cache["pdf1"]["parsed"] is parsed


def test_parse_error_sets_error_state(client):
    _doc_cache["bad1"] = _new_entry(
        path="C:/bad.png", name="bad.png", app="Image",
        parsed_state="discovered", source_type="image_file",
    )
    with patch("doc_intelligence.web.api.ImageParser") as MockImage:
        MockImage.return_value.parse_from_com.side_effect = RuntimeError("boom")
        resp = client.post("/api/documents/bad1/parse")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["parsed_state"] == "error"
    assert "boom" in (body.get("error") or "")
