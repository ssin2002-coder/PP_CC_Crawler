# Block 1: 문서 자동 감지 + 미리보기 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 열린 문서를 자동 감지하여 웹 브라우저에서 미리보기 + 파싱 데이터 + 핑거프린트 매칭을 표시하는 블록 1 구현

**Architecture:** Flask + SocketIO 백엔드가 COM 폴링으로 문서를 감지하고 파싱/핑거프린트 처리 후 WebSocket으로 React 프론트에 전달. 기존 com_worker/parsers/fingerprint/engine/storage 모듈을 그대로 재활용.

**Tech Stack:** Flask, flask-socketio (threading mode), pywin32, pyautogui, React 19, Vite, Zustand, socket.io-client

**Spec:** `docs/superpowers/specs/2026-05-10-block1-detect-preview-design.md`

---

## Task 0: 프로젝트 정리 — 불필요 파일 삭제

**Files:**
- Delete: `doc_intelligence/main.py`
- Delete: `doc_intelligence/ui_components.py`
- Delete: `doc_intelligence/anomaly.py`
- Delete: `doc_intelligence/graph.py`
- Delete: `doc_intelligence/region_linker.py`
- Delete: `doc_intelligence/validator.py`
- Delete: `doc_intelligence/extractor.py`
- Modify: `doc_intelligence/__init__.py`
- Delete: `run.py` (이전 tkinter 런처)

- [ ] **Step 1: 소스 파일 삭제**

```bash
git rm doc_intelligence/main.py doc_intelligence/ui_components.py doc_intelligence/anomaly.py doc_intelligence/graph.py doc_intelligence/region_linker.py doc_intelligence/validator.py doc_intelligence/extractor.py run.py
```

- [ ] **Step 2: 관련 테스트 파일 삭제**

삭제된 모듈을 임포트하는 테스트 파일도 함께 삭제해야 `pytest` 전체 실행 시 `ModuleNotFoundError`가 발생하지 않음.

```bash
git rm tests/test_anomaly.py tests/test_graph.py tests/test_region_linker.py tests/test_validator.py tests/test_extractor.py tests/test_ui.py tests/test_integration.py
```

- [ ] **Step 3: __init__.py 정리**

`doc_intelligence/__init__.py`를 빈 파일로 유지 (이미 비어있으면 스킵).

- [ ] **Step 4: 커밋**

```bash
git add -A
git commit -m "chore: remove tkinter UI, unused modules, and related tests for block 1 web migration"
```

---

## Task 1: 백엔드 — snapshot.py (윈도우 캡처)

**Files:**
- Create: `doc_intelligence/web/__init__.py`
- Create: `doc_intelligence/web/snapshot.py`
- Create: `tests/test_snapshot.py`

- [ ] **Step 1: 디렉토리 생성 및 __init__.py**

```bash
mkdir -p doc_intelligence/web
```

```python
# doc_intelligence/web/__init__.py
```

- [ ] **Step 2: 테스트 작성**

```python
# tests/test_snapshot.py
import pytest
from unittest.mock import patch, MagicMock
from doc_intelligence.web.snapshot import capture_window_snapshot


def test_capture_returns_base64_string():
    """캡처 결과가 base64 문자열인지 확인"""
    fake_img = MagicMock()
    import io
    buf = io.BytesIO()
    # 1x1 white PNG
    from PIL import Image
    Image.new("RGB", (1, 1), "white").save(buf, format="PNG")
    fake_img_bytes = buf.getvalue()

    with patch("doc_intelligence.web.snapshot.pyautogui") as mock_pyautogui:
        mock_screenshot = MagicMock()
        mock_screenshot.save = MagicMock(side_effect=lambda buf, **kw: buf.write(fake_img_bytes))
        mock_pyautogui.screenshot.return_value = mock_screenshot
        with patch("doc_intelligence.web.snapshot._get_window_rect", return_value=(0, 0, 100, 100)):
            result = capture_window_snapshot("test.xlsx")

    assert isinstance(result, str)
    assert len(result) > 0


def test_capture_returns_none_when_window_not_found():
    """창을 못 찾으면 None 반환"""
    with patch("doc_intelligence.web.snapshot._get_window_rect", return_value=None):
        result = capture_window_snapshot("nonexistent.xlsx")
    assert result is None
```

- [ ] **Step 3: 테스트 실행 — 실패 확인**

```bash
pytest tests/test_snapshot.py -v
```
Expected: FAIL (module not found)

- [ ] **Step 4: snapshot.py 구현**

```python
# doc_intelligence/web/snapshot.py
"""윈도우 캡처 — pyautogui 기반 스냅샷"""
import base64
import io
import logging

logger = logging.getLogger(__name__)

try:
    import pyautogui
    _PYAUTOGUI_AVAILABLE = True
except ImportError:
    pyautogui = None
    _PYAUTOGUI_AVAILABLE = False

try:
    import win32gui
    _WIN32GUI_AVAILABLE = True
except ImportError:
    win32gui = None
    _WIN32GUI_AVAILABLE = False


def _get_window_rect(filename: str):
    """파일명을 포함하는 윈도우의 (left, top, right, bottom) 반환. 없으면 None."""
    if not _WIN32GUI_AVAILABLE:
        return None

    result = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if filename in title:
                rect = win32gui.GetWindowRect(hwnd)
                result.append(rect)

    try:
        win32gui.EnumWindows(callback, None)
    except Exception:
        pass

    return result[0] if result else None


def capture_window_snapshot(filename: str) -> str | None:
    """파일명에 해당하는 윈도우를 캡처하여 base64 PNG 문자열로 반환.
    윈도우를 찾지 못하면 None.
    """
    if not _PYAUTOGUI_AVAILABLE:
        return None

    rect = _get_window_rect(filename)
    if rect is None:
        return None

    left, top, right, bottom = rect
    width = right - left
    height = bottom - top

    try:
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        buf = io.BytesIO()
        screenshot.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception as e:
        logger.warning("스냅샷 캡처 실패: %s", e)
        return None
```

- [ ] **Step 5: 테스트 실행 — 통과 확인**

```bash
pytest tests/test_snapshot.py -v
```
Expected: PASS

- [ ] **Step 6: 커밋**

```bash
git add doc_intelligence/web/__init__.py doc_intelligence/web/snapshot.py tests/test_snapshot.py
git commit -m "feat: add window snapshot capture module"
```

---

## Task 2: 백엔드 — app.py (Flask + SocketIO + COM 폴링)

**Files:**
- Create: `doc_intelligence/web/app.py`
- Create: `tests/test_app.py`

- [ ] **Step 1: 테스트 작성**

```python
# tests/test_app.py
import pytest
from unittest.mock import patch, MagicMock
from doc_intelligence.web.app import create_app, _build_doc_summary


def test_create_app_returns_flask_app():
    """create_app()이 Flask 앱을 반환하는지 확인"""
    app, socketio = create_app(testing=True)
    assert app is not None
    assert app.testing is True


def test_build_doc_summary_matched():
    """자동 매칭 문서의 summary 구조 확인"""
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
    assert result["score"] == 0.92


def test_build_doc_summary_candidate():
    """후보 문서의 summary 구조 확인"""
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
    """미매칭 문서의 summary 구조 확인"""
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
```

- [ ] **Step 2: 테스트 실행 — 실패 확인**

```bash
pytest tests/test_app.py -v
```
Expected: FAIL

- [ ] **Step 3: app.py 구현**

```python
# doc_intelligence/web/app.py
"""Flask + SocketIO 서버 + COM 폴링 스레드"""
import hashlib
import logging
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

# ProgID → 파서 매핑
_APP_TO_PARSER = {
    "Excel.Application": ExcelParser,
    "Word.Application": WordParser,
    "PowerPoint.Application": PowerPointParser,
    "AcroExch.App": PdfParser,
}

# 인메모리 캐시
_doc_cache: dict[str, dict] = {}
_polling_running = False


def _make_doc_id(file_path: str) -> str:
    return hashlib.md5(file_path.encode("utf-8")).hexdigest()


def _build_doc_summary(doc_id: str, entry: dict, template_names: dict) -> dict:
    """캐시 엔트리에서 프론트엔드용 요약 생성"""
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


def _polling_loop(com_worker: ComWorker, engine: Engine,
                  fingerprinter: Fingerprinter, socketio: SocketIO,
                  interval: int = 3):
    """COM 폴링 스레드 메인 루프"""
    global _polling_running
    _polling_running = True

    with com_worker.com_session():
        while _polling_running:
            try:
                docs = com_worker.detect_open_documents()
                current_ids = set()

                for doc_info in docs:
                    file_path = doc_info.get("path", "")
                    if not file_path:
                        continue

                    doc_id = _make_doc_id(file_path)
                    current_ids.add(doc_id)

                    # 이미 캐시에 있고 confirmed면 스킵
                    if doc_id in _doc_cache and _doc_cache[doc_id].get("confirmed"):
                        continue

                    # 이미 파싱된 문서면 스킵
                    if doc_id in _doc_cache and _doc_cache[doc_id].get("parsed"):
                        continue

                    # 새 문서 처리
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
                    except Exception as e:
                        logger.warning("문서 처리 실패 (%s): %s", file_path, e)

                # 닫힌 문서 제거
                closed = set(_doc_cache.keys()) - current_ids
                if closed:
                    for cid in closed:
                        del _doc_cache[cid]
                    socketio.emit("documents_updated", _get_all_summaries(engine))

            except Exception as e:
                logger.warning("폴링 루프 에러: %s", e)

            time.sleep(interval)


def _get_all_summaries(engine: Engine) -> list[dict]:
    """전체 문서 요약 목록 생성"""
    templates = engine.storage.get_all_templates()
    template_names = {t["id"]: t["name"] for t in templates}
    return [
        _build_doc_summary(doc_id, entry, template_names)
        for doc_id, entry in _doc_cache.items()
    ]


def create_app(testing=False, db_path="doc_intelligence.db"):
    """Flask 앱 팩토리"""
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

    # API 블루프린트 등록
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

    return app, socketio
```

- [ ] **Step 4: 테스트 실행 — 통과 확인**

```bash
pytest tests/test_app.py -v
```
Expected: PASS (api.py 미존재로 일부 실패 가능 → Task 3 후 재확인)

- [ ] **Step 5: 커밋**

```bash
git add doc_intelligence/web/app.py tests/test_app.py
git commit -m "feat: add Flask + SocketIO server with COM polling thread"
```

---

## Task 3: 백엔드 — api.py (REST 엔드포인트)

**Files:**
- Create: `doc_intelligence/web/api.py`
- Create: `tests/test_api.py`

- [ ] **Step 1: 테스트 작성**

```python
# tests/test_api.py
import pytest
from unittest.mock import MagicMock, patch
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
    """문서 없으면 빈 리스트"""
    resp = client.get("/api/documents")
    assert resp.status_code == 200
    assert resp.get_json() == []


def test_get_documents_with_cache(client):
    """캐시에 문서가 있으면 목록 반환"""
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
    """존재하지 않는 doc_id 요청 시 404"""
    resp = client.get("/api/documents/nonexistent/parsed")
    assert resp.status_code == 404


def test_get_parsed_success(client):
    """파싱 데이터 반환"""
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
    """학습 요청 시 템플릿 생성"""
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
    # 캐시가 갱신되었는지 확인
    assert _doc_cache["test123"]["match"]["auto"] is True


def test_get_templates(client):
    """템플릿 목록 조회"""
    resp = client.get("/api/templates")
    assert resp.status_code == 200
    assert isinstance(resp.get_json(), list)
```

- [ ] **Step 2: 테스트 실행 — 실패 확인**

```bash
pytest tests/test_api.py -v
```
Expected: FAIL

- [ ] **Step 3: api.py 구현**

```python
# doc_intelligence/web/api.py
"""REST API 엔드포인트"""
from dataclasses import asdict
from flask import Blueprint, jsonify, request

from doc_intelligence.engine import Engine
from doc_intelligence.fingerprint import Fingerprinter


def create_api_blueprint(engine: Engine, fingerprinter: Fingerprinter, doc_cache: dict):
    api = Blueprint("api", __name__)

    def _template_names() -> dict:
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
            "id": doc_id,
            "app": info.get("app", ""),
            "name": info.get("name", ""),
            "path": info.get("path", ""),
            "status": status,
            "score": round(score * 100, 1) if score else 0,
            "template_id": template_id,
            "template_name": names.get(template_id) if template_id else None,
            "labels": entry.get("fingerprint", {}).get("labels", []),
            "has_preview": entry.get("snapshot_b64") is not None,
        }

    # GET /api/documents
    @api.route("/documents", methods=["GET"])
    def get_documents():
        summaries = [_build_summary(did, e) for did, e in doc_cache.items()]
        return jsonify(summaries)

    # GET /api/documents/<id>/preview
    @api.route("/documents/<doc_id>/preview", methods=["GET"])
    def get_preview(doc_id):
        entry = doc_cache.get(doc_id)
        if entry is None:
            return jsonify({"error": "not found"}), 404
        b64 = entry.get("snapshot_b64")
        if b64 is None:
            return jsonify({"error": "no preview"}), 404
        return jsonify({"image": b64})

    # GET /api/documents/<id>/parsed
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
            "file_path": parsed.file_path,
            "file_type": parsed.file_type,
            "structure": parsed.structure,
            "cells": cells_data,
            "metadata": parsed.metadata,
        })

    # POST /api/templates/learn
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

        # 캐시 갱신 (confirmed는 confirm 엔드포인트에서만 설정)
        entry["match"] = {"template": template_id, "score": 1.0, "auto": True}

        return jsonify({"template_id": template_id, "status": "learned"})

    # POST /api/templates/confirm
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

    # GET /api/templates
    @api.route("/templates", methods=["GET"])
    def get_templates():
        templates = engine.storage.get_all_templates()
        return jsonify(templates)

    return api
```

- [ ] **Step 4: 테스트 실행 — 통과 확인**

```bash
pytest tests/test_api.py tests/test_app.py -v
```
Expected: PASS

- [ ] **Step 5: 커밋**

```bash
git add doc_intelligence/web/api.py tests/test_api.py
git commit -m "feat: add REST API endpoints for documents and templates"
```

---

## Task 4: 백엔드 — run.py (진입점)

**Files:**
- Create: `run.py`

- [ ] **Step 1: run.py 작성**

```python
# run.py
"""Doc Intelligence 웹 서버 진입점"""
import webbrowser
import threading

from doc_intelligence.web.app import create_app


def main():
    app, socketio = create_app()
    port = 5000

    # 1초 후 브라우저 오픈
    threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{port}")).start()

    print(f"Doc Intelligence running at http://localhost:{port}")
    socketio.run(app, host="0.0.0.0", port=port, debug=False)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 커밋**

```bash
git add run.py
git commit -m "feat: add web server entry point with browser auto-open"
```

---

## Task 5: 프론트엔드 — React 프로젝트 초기화

**Files:**
- Create: `doc_intelligence/web/frontend/package.json`
- Create: `doc_intelligence/web/frontend/vite.config.js`
- Create: `doc_intelligence/web/frontend/index.html`
- Create: `doc_intelligence/web/frontend/src/main.jsx`
- Create: `doc_intelligence/web/frontend/src/App.jsx`
- Create: `doc_intelligence/web/frontend/src/index.css`

- [ ] **Step 1: Vite + React 프로젝트 생성 및 package.json 설정**

```bash
cd doc_intelligence/web/frontend
npm init -y
npm install react react-dom zustand socket.io-client
npm install -D vite @vitejs/plugin-react
```

생성된 `package.json`의 `scripts` 섹션을 다음으로 수정:

```json
{
  "scripts": {
    "dev": "vite",
    "build": "vite build",
    "preview": "vite preview"
  }
}
```

- [ ] **Step 2: vite.config.js 작성**

```javascript
// doc_intelligence/web/frontend/vite.config.js
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/api': 'http://localhost:5000',
      '/socket.io': {
        target: 'http://localhost:5000',
        ws: true,
      },
    },
  },
  build: {
    outDir: '../static',
    emptyOutDir: true,
  },
});
```

- [ ] **Step 3: index.html 작성**

```html
<!-- doc_intelligence/web/frontend/index.html -->
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Doc Intelligence</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet" />
</head>
<body>
  <div id="root"></div>
  <script type="module" src="/src/main.jsx"></script>
</body>
</html>
```

- [ ] **Step 4: index.css — Apple Dark 테마 토큰**

```css
/* doc_intelligence/web/frontend/src/index.css */
:root {
  --bg-main: #000000;
  --bg-panel: #1d1d1f;
  --bg-card: #1d1d1f;
  --border: #333336;
  --text-main: #f5f5f7;
  --text-sub: #86868b;
  --accent-blue: #0071e3;
  --accent-blue-light: #2997ff;
  --color-green: #30d158;
  --color-orange: #ff9f0a;
  --radius-card: 12px;
  --radius-pill: 980px;
  --font-family: 'Inter', system-ui, -apple-system, sans-serif;
}

* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: var(--font-family);
  background: var(--bg-main);
  color: var(--text-main);
  -webkit-font-smoothing: antialiased;
}

::-webkit-scrollbar {
  width: 6px;
}

::-webkit-scrollbar-track {
  background: var(--bg-main);
}

::-webkit-scrollbar-thumb {
  background: var(--border);
  border-radius: 3px;
}
```

- [ ] **Step 5: main.jsx + App.jsx 스켈레톤**

```jsx
// doc_intelligence/web/frontend/src/main.jsx
import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
```

```jsx
// doc_intelligence/web/frontend/src/App.jsx
import TopBar from './components/TopBar';
import FileList from './components/FileList';
import FingerInfo from './components/FingerInfo';
import DataTable from './components/DataTable';
import { useStore } from './stores/store';

export default function App() {
  const selectedDocId = useStore((s) => s.selectedDocId);

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <TopBar />
      <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
        <div style={{ width: '33.3%', borderRight: '1px solid var(--border)', overflow: 'auto' }}>
          <FileList />
        </div>
        <div style={{ width: '66.7%', display: 'flex', flexDirection: 'column', overflow: 'auto' }}>
          {selectedDocId ? (
            <>
              <FingerInfo />
              <DataTable />
            </>
          ) : (
            <div style={{
              flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center',
              color: 'var(--text-sub)', fontSize: '14px'
            }}>
              좌측에서 문서를 선택하세요
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
```

- [ ] **Step 6: 커밋**

```bash
git add doc_intelligence/web/frontend/
git commit -m "feat: initialize React + Vite frontend with Apple Dark theme"
```

---

## Task 6: 프론트엔드 — Zustand 스토어 + WebSocket 훅

**Files:**
- Create: `doc_intelligence/web/frontend/src/stores/store.js`
- Create: `doc_intelligence/web/frontend/src/hooks/useSocket.js`

- [ ] **Step 1: Zustand 스토어**

```javascript
// doc_intelligence/web/frontend/src/stores/store.js
import { create } from 'zustand';

export const useStore = create((set, get) => ({
  documents: [],
  selectedDocId: null,
  parsedData: {},
  comStatus: 'connecting',

  setDocuments: (docs) => set({ documents: docs }),
  selectDocument: (docId) => set({ selectedDocId: docId }),
  setComStatus: (status) => set({ comStatus: status }),

  setParsedData: (docId, data) =>
    set((state) => ({
      parsedData: { ...state.parsedData, [docId]: data },
    })),

  fetchParsed: async (docId) => {
    if (get().parsedData[docId]) return;
    try {
      const res = await fetch(`/api/documents/${docId}/parsed`);
      if (res.ok) {
        const data = await res.json();
        get().setParsedData(docId, data);
      }
    } catch (e) {
      console.error('Failed to fetch parsed data:', e);
    }
  },
}));
```

- [ ] **Step 2: WebSocket 훅**

```javascript
// doc_intelligence/web/frontend/src/hooks/useSocket.js
import { useEffect } from 'react';
import { io } from 'socket.io-client';
import { useStore } from '../stores/store';

let socket = null;

export function useSocket() {
  const setDocuments = useStore((s) => s.setDocuments);
  const setComStatus = useStore((s) => s.setComStatus);

  useEffect(() => {
    socket = io({ transports: ['websocket', 'polling'] });

    socket.on('connect', () => setComStatus('connected'));
    socket.on('disconnect', () => setComStatus('disconnected'));
    socket.on('documents_updated', (docs) => setDocuments(docs));
    socket.on('parse_complete', ({ doc_id }) => {
      // 파싱 완료 시 해당 문서가 선택되어 있으면 자동으로 파싱 데이터 fetch
      const state = useStore.getState();
      if (state.selectedDocId === doc_id) {
        state.fetchParsed(doc_id);
      }
    });

    return () => {
      socket.disconnect();
      socket = null;
    };
  }, [setDocuments, setComStatus]);
}
```

- [ ] **Step 3: 커밋**

```bash
git add doc_intelligence/web/frontend/src/stores/ doc_intelligence/web/frontend/src/hooks/
git commit -m "feat: add Zustand store and WebSocket hook"
```

---

## Task 7: 프론트엔드 — TopBar 컴포넌트

**Files:**
- Create: `doc_intelligence/web/frontend/src/components/TopBar.jsx`

- [ ] **Step 1: TopBar.jsx**

```jsx
// doc_intelligence/web/frontend/src/components/TopBar.jsx
import { useStore } from '../stores/store';

const styles = {
  bar: {
    display: 'flex', justifyContent: 'space-between', alignItems: 'center',
    background: 'var(--bg-panel)', padding: '10px 20px',
    borderBottom: '1px solid var(--border)', flexShrink: 0,
  },
  left: { display: 'flex', alignItems: 'center', gap: '8px' },
  title: { fontSize: '15px', fontWeight: 600, color: '#fff', letterSpacing: '-0.015em' },
  version: { fontSize: '11px', color: 'var(--text-sub)' },
  right: { display: 'flex', alignItems: 'center', gap: '12px' },
  dot: (connected) => ({
    width: '8px', height: '8px', borderRadius: '50%',
    background: connected ? 'var(--color-green)' : '#ff453a',
  }),
  status: { fontSize: '12px', color: 'var(--text-sub)' },
  btnPrimary: {
    background: 'var(--accent-blue)', color: '#fff', border: 'none',
    borderRadius: 'var(--radius-pill)', padding: '5px 14px',
    fontSize: '12px', cursor: 'pointer', fontWeight: 500,
  },
  btnNormal: {
    background: 'var(--bg-card)', color: 'var(--text-main)', border: '1px solid var(--border)',
    borderRadius: 'var(--radius-pill)', padding: '5px 14px',
    fontSize: '12px', cursor: 'pointer',
  },
};

export default function TopBar() {
  const documents = useStore((s) => s.documents);
  const comStatus = useStore((s) => s.comStatus);
  const connected = comStatus === 'connected';

  return (
    <div style={styles.bar}>
      <div style={styles.left}>
        <span style={styles.title}>Doc Intelligence</span>
        <span style={styles.version}>v0.2</span>
      </div>
      <div style={styles.right}>
        <span style={styles.dot(connected)} />
        <span style={styles.status}>
          {connected ? 'COM 연결됨' : 'COM 연결 끊김'} | 문서 {documents.length}개 열림
        </span>
        <button style={styles.btnPrimary}>+ 영역 연결</button>
        <button style={styles.btnNormal}>설정</button>
      </div>
    </div>
  );
}
```

- [ ] **Step 2: 커밋**

```bash
git add doc_intelligence/web/frontend/src/components/TopBar.jsx
git commit -m "feat: add TopBar component with Apple Dark theme"
```

---

## Task 8: 프론트엔드 — FileList + FileCard 컴포넌트

**Files:**
- Create: `doc_intelligence/web/frontend/src/components/FileList.jsx`
- Create: `doc_intelligence/web/frontend/src/components/FileCard.jsx`
- Create: `doc_intelligence/web/frontend/src/components/DocPreview.jsx`

- [ ] **Step 1: DocPreview.jsx**

```jsx
// doc_intelligence/web/frontend/src/components/DocPreview.jsx

export default function DocPreview({ docId, hasPreview }) {
  if (!hasPreview) {
    return (
      <div style={{
        background: '#000', borderRadius: '8px', height: '80px',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        border: '1px solid var(--border)', fontSize: '11px', color: 'var(--text-sub)',
      }}>
        미리보기 없음
      </div>
    );
  }

  return (
    <img
      src={`/api/documents/${docId}/preview`}
      alt="preview"
      style={{
        width: '100%', height: '80px', objectFit: 'cover',
        borderRadius: '8px', border: '1px solid var(--border)',
      }}
      onError={(e) => { e.target.style.display = 'none'; }}
    />
  );
}
```

- [ ] **Step 2: FileCard.jsx**

```jsx
// doc_intelligence/web/frontend/src/components/FileCard.jsx
import DocPreview from './DocPreview';

const iconMap = {
  'Excel.Application': '📊',
  'Word.Application': '📝',
  'PowerPoint.Application': '📑',
};

const statusConfig = {
  matched: { label: '매칭됨', bg: 'var(--color-green)', color: '#000' },
  candidate: { label: '후보', bg: 'var(--color-orange)', color: '#000' },
  new: { label: '새 문서', bg: 'var(--accent-blue-light)', color: '#fff' },
};

export default function FileCard({ doc, selected, onClick }) {
  const icon = iconMap[doc.app] || '📄';
  const badge = statusConfig[doc.status] || statusConfig.new;

  return (
    <div
      onClick={onClick}
      style={{
        background: selected ? 'rgba(0, 113, 227, 0.15)' : 'var(--bg-card)',
        border: `1px solid ${selected ? 'var(--accent-blue)' : 'var(--border)'}`,
        borderRadius: 'var(--radius-card)', padding: '10px',
        marginBottom: '8px', cursor: 'pointer',
        transition: 'border-color 0.2s',
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' }}>
        <span style={{ fontSize: '16px' }}>{icon}</span>
        <span style={{
          fontSize: '12px', color: 'var(--text-main)', fontWeight: 500,
          overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
        }}>
          {doc.name}
        </span>
      </div>
      <DocPreview docId={doc.id} hasPreview={doc.has_preview} />
      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '6px' }}>
        <span style={{
          fontSize: '10px', background: badge.bg, color: badge.color,
          padding: '1px 8px', borderRadius: 'var(--radius-pill)', fontWeight: 500,
        }}>
          {badge.label}
        </span>
        {doc.template_name && (
          <span style={{ fontSize: '10px', color: 'var(--text-sub)' }}>
            {doc.template_name} ({doc.score}%)
          </span>
        )}
      </div>
    </div>
  );
}
```

- [ ] **Step 3: FileList.jsx**

```jsx
// doc_intelligence/web/frontend/src/components/FileList.jsx
import { useSocket } from '../hooks/useSocket';
import { useStore } from '../stores/store';
import FileCard from './FileCard';

export default function FileList() {
  useSocket();

  const documents = useStore((s) => s.documents);
  const selectedDocId = useStore((s) => s.selectedDocId);
  const selectDocument = useStore((s) => s.selectDocument);
  const fetchParsed = useStore((s) => s.fetchParsed);

  const handleSelect = (docId) => {
    selectDocument(docId);
    fetchParsed(docId);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div style={{
        background: 'var(--bg-panel)', padding: '10px 14px',
        display: 'flex', justifyContent: 'space-between', alignItems: 'center',
        borderBottom: '1px solid var(--border)',
      }}>
        <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-main)' }}>
          열린 문서
        </span>
        <span style={{ fontSize: '12px', color: 'var(--accent-blue-light)' }}>
          {documents.length}개 감지
        </span>
      </div>
      <div style={{ flex: 1, overflow: 'auto', padding: '8px' }}>
        {documents.length === 0 ? (
          <div style={{
            textAlign: 'center', color: 'var(--text-sub)',
            fontSize: '12px', marginTop: '40px',
          }}>
            열린 문서가 없습니다
          </div>
        ) : (
          documents.map((doc) => (
            <FileCard
              key={doc.id}
              doc={doc}
              selected={doc.id === selectedDocId}
              onClick={() => handleSelect(doc.id)}
            />
          ))
        )}
      </div>
    </div>
  );
}
```

- [ ] **Step 4: 커밋**

```bash
git add doc_intelligence/web/frontend/src/components/FileList.jsx doc_intelligence/web/frontend/src/components/FileCard.jsx doc_intelligence/web/frontend/src/components/DocPreview.jsx
git commit -m "feat: add FileList, FileCard, DocPreview components"
```

---

## Task 9: 프론트엔드 — FingerInfo 컴포넌트

**Files:**
- Create: `doc_intelligence/web/frontend/src/components/FingerInfo.jsx`

- [ ] **Step 1: FingerInfo.jsx**

```jsx
// doc_intelligence/web/frontend/src/components/FingerInfo.jsx
import { useState } from 'react';
import { useStore } from '../stores/store';

export default function FingerInfo() {
  const selectedDocId = useStore((s) => s.selectedDocId);
  const documents = useStore((s) => s.documents);
  const [learnName, setLearnName] = useState('');
  const [showModal, setShowModal] = useState(false);

  const doc = documents.find((d) => d.id === selectedDocId);
  if (!doc) return null;

  const handleLearn = async () => {
    if (!learnName.trim()) return;
    const res = await fetch('/api/templates/learn', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ doc_id: doc.id, template_name: learnName }),
    });
    if (res.ok) {
      setShowModal(false);
      setLearnName('');
    }
  };

  const handleConfirm = async () => {
    await fetch('/api/templates/confirm', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ doc_id: doc.id, template_id: doc.template_id }),
    });
  };

  return (
    <div style={{
      background: 'var(--bg-panel)', borderBottom: '1px solid var(--border)',
      padding: '12px 16px',
    }}>
      {/* 매칭 상태 */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
        {doc.status === 'matched' && (
          <span style={{ fontSize: '13px', color: 'var(--color-green)' }}>
            ✓ {doc.template_name} ({doc.score}%)
          </span>
        )}
        {doc.status === 'candidate' && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '13px', color: 'var(--color-orange)' }}>
              ? 이 양식인가요: {doc.template_name} ({doc.score}%)
            </span>
            <button onClick={handleConfirm} style={{
              background: 'var(--color-green)', color: '#000', border: 'none',
              borderRadius: 'var(--radius-pill)', padding: '3px 12px',
              fontSize: '11px', cursor: 'pointer', fontWeight: 500,
            }}>예</button>
            <button onClick={() => setShowModal(true)} style={{
              background: 'var(--bg-card)', color: 'var(--text-main)',
              border: '1px solid var(--border)',
              borderRadius: 'var(--radius-pill)', padding: '3px 12px',
              fontSize: '11px', cursor: 'pointer',
            }}>아니오</button>
          </div>
        )}
        {doc.status === 'new' && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '13px', color: 'var(--accent-blue-light)' }}>
              새 문서 — 템플릿 학습 필요
            </span>
            <button onClick={() => setShowModal(true)} style={{
              background: 'var(--accent-blue)', color: '#fff', border: 'none',
              borderRadius: 'var(--radius-pill)', padding: '3px 12px',
              fontSize: '11px', cursor: 'pointer', fontWeight: 500,
            }}>학습</button>
          </div>
        )}
      </div>

      {/* 필드 태그 */}
      {doc.labels && doc.labels.length > 0 && (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px' }}>
          {doc.labels.slice(0, 20).map((label, i) => (
            <span key={i} style={{
              fontSize: '10px', background: 'rgba(255,255,255,0.08)',
              color: 'var(--text-sub)', padding: '2px 8px',
              borderRadius: 'var(--radius-pill)', border: '1px solid var(--border)',
            }}>
              {label}
            </span>
          ))}
          {doc.labels.length > 20 && (
            <span style={{ fontSize: '10px', color: 'var(--text-sub)' }}>
              +{doc.labels.length - 20}
            </span>
          )}
        </div>
      )}

      {/* 학습 모달 */}
      {showModal && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.7)',
          display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000,
        }}>
          <div style={{
            background: 'var(--bg-panel)', borderRadius: 'var(--radius-card)',
            padding: '24px', width: '360px', border: '1px solid var(--border)',
          }}>
            <h3 style={{ fontSize: '15px', marginBottom: '12px', fontWeight: 600 }}>
              템플릿 학습
            </h3>
            <input
              value={learnName}
              onChange={(e) => setLearnName(e.target.value)}
              placeholder="양식 이름 입력 (예: 정비비용정산서)"
              style={{
                width: '100%', padding: '8px 12px', borderRadius: '8px',
                border: '1px solid var(--border)', background: '#000',
                color: 'var(--text-main)', fontSize: '13px', outline: 'none',
              }}
              onKeyDown={(e) => e.key === 'Enter' && handleLearn()}
              autoFocus
            />
            <div style={{ display: 'flex', gap: '8px', marginTop: '16px', justifyContent: 'flex-end' }}>
              <button onClick={() => setShowModal(false)} style={{
                background: 'var(--bg-card)', color: 'var(--text-main)',
                border: '1px solid var(--border)',
                borderRadius: 'var(--radius-pill)', padding: '6px 16px',
                fontSize: '12px', cursor: 'pointer',
              }}>취소</button>
              <button onClick={handleLearn} style={{
                background: 'var(--accent-blue)', color: '#fff', border: 'none',
                borderRadius: 'var(--radius-pill)', padding: '6px 16px',
                fontSize: '12px', cursor: 'pointer', fontWeight: 500,
              }}>학습</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
```

- [ ] **Step 2: 커밋**

```bash
git add doc_intelligence/web/frontend/src/components/FingerInfo.jsx
git commit -m "feat: add FingerInfo component with learn/confirm UI"
```

---

## Task 10: 프론트엔드 — DataTable 컴포넌트

**Files:**
- Create: `doc_intelligence/web/frontend/src/components/DataTable.jsx`

- [ ] **Step 1: DataTable.jsx**

```jsx
// doc_intelligence/web/frontend/src/components/DataTable.jsx
import { useState } from 'react';
import { useStore } from '../stores/store';

export default function DataTable() {
  const selectedDocId = useStore((s) => s.selectedDocId);
  const parsedData = useStore((s) => s.parsedData[selectedDocId]);
  const [activeSheet, setActiveSheet] = useState(null);

  if (!parsedData) {
    return (
      <div style={{
        flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center',
        color: 'var(--text-sub)', fontSize: '13px',
      }}>
        파싱 데이터 로딩 중...
      </div>
    );
  }

  // 셀을 시트별로 그룹핑
  const cells = parsedData.cells || [];
  const sheets = {};
  cells.forEach((cell) => {
    // address 형식: "Sheet1!R1C1" 또는 "para:1"
    const parts = cell.address.split('!');
    const sheetName = parts.length > 1 ? parts[0] : '_default';
    if (!sheets[sheetName]) sheets[sheetName] = [];
    sheets[sheetName].push(cell);
  });

  const sheetNames = Object.keys(sheets);
  const currentSheet = activeSheet || sheetNames[0] || '_default';
  const sheetCells = sheets[currentSheet] || [];

  // 셀을 2D 그리드로 변환
  const grid = {};
  let maxRow = 0;
  let maxCol = 0;
  sheetCells.forEach((cell) => {
    const match = cell.address.match(/R(\d+)C(\d+)/);
    if (match) {
      const r = parseInt(match[1]);
      const c = parseInt(match[2]);
      if (!grid[r]) grid[r] = {};
      grid[r][c] = cell.value;
      maxRow = Math.max(maxRow, r);
      maxCol = Math.max(maxCol, c);
    }
  });

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      {/* 시트 탭 */}
      {sheetNames.length > 1 && (
        <div style={{
          display: 'flex', gap: '0', borderBottom: '1px solid var(--border)',
          background: 'var(--bg-panel)', flexShrink: 0,
        }}>
          {sheetNames.map((name) => (
            <button
              key={name}
              onClick={() => setActiveSheet(name)}
              style={{
                padding: '6px 16px', fontSize: '11px', cursor: 'pointer',
                border: 'none', borderBottom: name === currentSheet ? '2px solid var(--accent-blue)' : '2px solid transparent',
                background: 'transparent',
                color: name === currentSheet ? 'var(--text-main)' : 'var(--text-sub)',
                fontWeight: name === currentSheet ? 500 : 400,
              }}
            >
              {name}
            </button>
          ))}
        </div>
      )}

      {/* 테이블 */}
      <div style={{ flex: 1, overflow: 'auto', padding: '8px' }}>
        {maxRow === 0 ? (
          <div style={{ color: 'var(--text-sub)', fontSize: '12px', textAlign: 'center', marginTop: '20px' }}>
            표시할 데이터가 없습니다
          </div>
        ) : (
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px' }}>
            <tbody>
              {Array.from({ length: maxRow }, (_, ri) => ri + 1).map((r) => (
                <tr key={r}>
                  {Array.from({ length: maxCol }, (_, ci) => ci + 1).map((c) => (
                    <td key={c} style={{
                      border: '1px solid var(--border)', padding: '4px 6px',
                      color: 'var(--text-main)', whiteSpace: 'nowrap',
                      background: r === 1 ? 'var(--bg-panel)' : 'transparent',
                      fontWeight: r === 1 ? 500 : 400,
                    }}>
                      {grid[r]?.[c] ?? ''}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}
```

- [ ] **Step 2: 커밋**

```bash
git add doc_intelligence/web/frontend/src/components/DataTable.jsx
git commit -m "feat: add DataTable component with sheet tabs and grid rendering"
```

---

## Task 11: 통합 — Flask에서 React 빌드 제공

**Files:**
- Modify: `doc_intelligence/web/app.py`

- [ ] **Step 1: app.py에 static 파일 서빙 추가**

`create_app()` 함수에서 빌드된 React 파일을 서빙하도록 수정:

```python
# create_app() 내부에 추가:
import os

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
```

- [ ] **Step 2: 프론트엔드 빌드 테스트**

```bash
cd doc_intelligence/web/frontend
npm run build
```
Expected: `doc_intelligence/web/static/` 디렉토리에 빌드 파일 생성

- [ ] **Step 3: 커밋**

```bash
git add doc_intelligence/web/app.py doc_intelligence/web/static/
git commit -m "feat: serve React build from Flask and integrate frontend"
```

---

## Task 12: E2E 검증

- [ ] **Step 1: 백엔드 전체 테스트 실행**

```bash
pytest tests/test_snapshot.py tests/test_app.py tests/test_api.py -v
```
Expected: 전체 PASS

- [ ] **Step 2: 수동 E2E 테스트**

1. Excel 파일 하나 열기
2. `python run.py` 실행
3. 브라우저에서 `http://localhost:5000` 접속
4. 확인 사항:
   - 좌측 패널에 열린 문서가 자동 감지되어 표시
   - 파일 카드에 스냅샷 미리보기 또는 "미리보기 없음" 표시
   - 상태 뱃지 표시 (새 문서 / 후보 / 매칭됨)
   - 파일 카드 클릭 시 우측에 핑거프린트 정보 + 파싱 테이블 표시
   - [학습] 버튼 클릭 → 모달 → 이름 입력 → 학습 완료 → 뱃지 전환

- [ ] **Step 3: 최종 커밋**

```bash
git add -A
git commit -m "feat: Block 1 complete — document auto-detection with web UI"
```
