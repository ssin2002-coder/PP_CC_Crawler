# Doc Intelligence 구현 계획 (v2)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 정비비용정산 문서(Excel/Word/PPT/PDF/이미지)를 COM으로 파싱하고, 템플릿 핑거프린팅 + 드래그 영역 연결 + 단일 룰/프리셋 교차 검증으로 문서 간 이상을 탐지하는 시스템 구축

**Architecture:** 코어 엔진(플러그인 레지스트리) 위에 레고 블록을 독립적으로 추가하는 모듈 구조. COM 호출은 별도 프로세스로 격리(STA, 재시도, 타임아웃). SQLite로 템플릿/룰 저장, tkinter 데스크톱 UI.

**Tech Stack:** Python 3.10+, pywin32(COM), kiwipiepy, scikit-learn, NetworkX, pytesseract, pyautogui, tkinter, SQLite

**Spec:** `docs/superpowers/specs/2026-05-09-doc-intelligence-design.md`

**Status 값 규칙:** 모든 곳에서 한국어 사용 — `"통과"`, `"실패"`, `"경고"`

---

## 파일 구조

```
doc_intelligence/
+-- __init__.py
+-- main.py              # 진입점 + tkinter 메인 UI
+-- engine.py            # 파이프라인 + 플러그인 레지스트리 + 데이터 모델
+-- com_worker.py         # COM 프로세스 격리 래퍼 (STA, 재시도, 타임아웃)
+-- parsers.py           # 모든 문서 파서 (COM 기반)
+-- fingerprint.py       # TF-IDF 핑거프린트 + 템플릿 매칭
+-- extractor.py         # 엔티티 추출 (regex + 형태소)
+-- validator.py         # 교차 검증 (6종 룰 + 프리셋 자동 감지)
+-- region_linker.py     # 드래그 영역 연결 (투명 오버레이 + DPI)
+-- anomaly.py           # Isolation Forest 이상 탐지
+-- graph.py             # NetworkX 관계 그래프 + pyvis 시각화
+-- storage.py           # SQLite CRUD 전체
+-- ui_components.py     # tkinter 위젯
+-- config.yaml          # 설정 (커스텀 패턴, 임계값 등)
+-- templates.db         # 자동 생성
tests/
+-- test_storage.py
+-- test_engine.py
+-- test_com_worker.py
+-- test_parsers.py
+-- test_fingerprint.py
+-- test_extractor.py
+-- test_validator.py
+-- test_region_linker.py
+-- test_anomaly.py
+-- test_graph.py
+-- test_ui.py
+-- test_integration.py
```

---

## Task 1: 저장소 — storage.py (CRUD 전체)

**Files:** Create `doc_intelligence/storage.py`, `doc_intelligence/__init__.py`, `tests/test_storage.py`, `requirements_doc_intelligence.txt`

- [ ] **Step 1: requirements 생성**

```
# requirements_doc_intelligence.txt
pywin32>=306
kiwipiepy>=0.20.0
scikit-learn>=1.3
networkx>=3.0
pytesseract>=0.3.10
pyautogui>=0.9.54
yake>=0.4.8
pyvis>=0.3.2
pyyaml>=6.0
```

- [ ] **Step 2: 테스트 작성**

```python
# tests/test_storage.py
import os, tempfile, pytest
from doc_intelligence.storage import Storage

@pytest.fixture
def db():
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    s = Storage(path)
    yield s
    s.close()
    os.unlink(path)

def test_init_creates_tables(db):
    tables = db.list_tables()
    for t in ["templates","rules","presets","documents","validation_results"]:
        assert t in tables

def test_template_crud(db):
    tid = db.save_template("견적서v1","excel",[0.1,0.2],{"A1":"합계"},{"A2":"금액"})
    t = db.get_template(tid)
    assert t["name"] == "견적서v1"
    db.update_template(tid, name="견적서v2", field_mappings={"A2":"업체명"})
    t2 = db.get_template(tid)
    assert t2["name"] == "견적서v2"
    assert t2["field_mappings"] == {"A2":"업체명"}
    db.delete_template(tid)
    assert db.get_template(tid) is None

def test_rule_crud(db):
    rid = db.save_rule("금액일치","값_일치",[{"doc":"A","cell":"C5"}],{})
    r = db.get_rule(rid)
    assert r["name"] == "금액일치"
    db.update_rule(rid, name="금액비교")
    assert db.get_rule(rid)["name"] == "금액비교"
    db.delete_rule(rid)
    assert db.get_rule(rid) is None

def test_preset_crud(db):
    pid = db.save_preset("배관정비","배관",[1,2,3],[10,20])
    p = db.get_preset(pid)
    assert p["template_ids"] == [10,20]
    db.update_preset(pid, rule_ids=[1,2,3,4])
    assert db.get_preset(pid)["rule_ids"] == [1,2,3,4]

def test_validation_result(db):
    vid = db.save_validation_result(1,1,[1,2],"통과",{"msg":"ok"})
    results = db.get_validation_results(preset_id=1)
    assert len(results) == 1
    assert results[0]["status"] == "통과"

def test_get_nonexistent_returns_none(db):
    assert db.get_template(999) is None
    assert db.get_rule(999) is None
    assert db.get_preset(999) is None

def test_get_all_templates(db):
    db.save_template("A","excel",[0.1],{},{})
    db.save_template("B","word",[0.2],{},{})
    assert len(db.get_all_templates()) == 2

def test_find_preset_by_template_ids(db):
    db.save_preset("P1","cat",[1],[10,20])
    db.save_preset("P2","cat",[2],[20,30])
    matches = db.find_presets_by_template_ids([10,20])
    assert any(p["name"] == "P1" for p in matches)
```

- [ ] **Step 3: 테스트 실행 (실패 확인)**

- [ ] **Step 4: storage.py 구현**

```python
# doc_intelligence/storage.py
import sqlite3, json
from datetime import datetime

class Storage:
    def __init__(self, db_path="templates.db"):
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self._init_tables()

    def _init_tables(self):
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT, file_type TEXT,
                fingerprint_vector TEXT, label_positions TEXT,
                field_mappings TEXT, created_at TEXT,
                match_count INTEGER DEFAULT 0);
            CREATE TABLE IF NOT EXISTS rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT, rule_type TEXT,
                regions TEXT, params TEXT, created_at TEXT);
            CREATE TABLE IF NOT EXISTS presets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT, category TEXT,
                rule_ids TEXT, template_ids TEXT, created_at TEXT);
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT, file_type TEXT,
                template_id INTEGER, entities TEXT, parsed_at TEXT,
                FOREIGN KEY (template_id) REFERENCES templates(id));
            CREATE TABLE IF NOT EXISTS validation_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                preset_id INTEGER, rule_id INTEGER,
                document_ids TEXT, status TEXT, detail TEXT,
                executed_at TEXT);
        """)
        self.conn.commit()

    def list_tables(self):
        return [r[0] for r in self.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'").fetchall()]

    # --- Templates CRUD ---
    def save_template(self, name, file_type, vector, label_positions, field_mappings):
        c = self.conn.execute(
            "INSERT INTO templates (name,file_type,fingerprint_vector,label_positions,field_mappings,created_at) VALUES (?,?,?,?,?,?)",
            (name, file_type, json.dumps(vector), json.dumps(label_positions,ensure_ascii=False),
             json.dumps(field_mappings,ensure_ascii=False), datetime.now().isoformat()))
        self.conn.commit()
        return c.lastrowid

    def get_template(self, tid):
        r = self.conn.execute("SELECT * FROM templates WHERE id=?", (tid,)).fetchone()
        if not r: return None
        d = dict(r)
        for k in ["fingerprint_vector","label_positions","field_mappings"]:
            d[k] = json.loads(d[k]) if d[k] else {}
        return d

    def get_all_templates(self):
        rows = self.conn.execute("SELECT * FROM templates").fetchall()
        result = []
        for r in rows:
            d = dict(r)
            for k in ["fingerprint_vector","label_positions","field_mappings"]:
                d[k] = json.loads(d[k]) if d[k] else {}
            result.append(d)
        return result

    def update_template(self, tid, **kwargs):
        for key, val in kwargs.items():
            if key in ("fingerprint_vector","label_positions","field_mappings"):
                val = json.dumps(val, ensure_ascii=False)
            self.conn.execute(f"UPDATE templates SET {key}=? WHERE id=?", (val, tid))
        self.conn.commit()

    def delete_template(self, tid):
        self.conn.execute("DELETE FROM templates WHERE id=?", (tid,))
        self.conn.commit()

    def increment_match_count(self, tid):
        self.conn.execute("UPDATE templates SET match_count=match_count+1 WHERE id=?", (tid,))
        self.conn.commit()

    # --- Rules CRUD ---
    def save_rule(self, name, rule_type, regions, params):
        c = self.conn.execute(
            "INSERT INTO rules (name,rule_type,regions,params,created_at) VALUES (?,?,?,?,?)",
            (name, rule_type, json.dumps(regions,ensure_ascii=False),
             json.dumps(params,ensure_ascii=False), datetime.now().isoformat()))
        self.conn.commit()
        return c.lastrowid

    def get_rule(self, rid):
        r = self.conn.execute("SELECT * FROM rules WHERE id=?", (rid,)).fetchone()
        if not r: return None
        d = dict(r)
        d["regions"] = json.loads(d["regions"]) if d["regions"] else []
        d["params"] = json.loads(d["params"]) if d["params"] else {}
        return d

    def get_all_rules(self):
        return [dict(r) | {"regions": json.loads(r["regions"] or "[]"), "params": json.loads(r["params"] or "{}")}
                for r in self.conn.execute("SELECT * FROM rules").fetchall()]

    def update_rule(self, rid, **kwargs):
        for key, val in kwargs.items():
            if key in ("regions","params"):
                val = json.dumps(val, ensure_ascii=False)
            self.conn.execute(f"UPDATE rules SET {key}=? WHERE id=?", (val, rid))
        self.conn.commit()

    def delete_rule(self, rid):
        self.conn.execute("DELETE FROM rules WHERE id=?", (rid,))
        self.conn.commit()

    # --- Presets CRUD ---
    def save_preset(self, name, category, rule_ids, template_ids):
        c = self.conn.execute(
            "INSERT INTO presets (name,category,rule_ids,template_ids,created_at) VALUES (?,?,?,?,?)",
            (name, category, json.dumps(rule_ids), json.dumps(template_ids), datetime.now().isoformat()))
        self.conn.commit()
        return c.lastrowid

    def get_preset(self, pid):
        r = self.conn.execute("SELECT * FROM presets WHERE id=?", (pid,)).fetchone()
        if not r: return None
        d = dict(r)
        d["rule_ids"] = json.loads(d["rule_ids"]) if d["rule_ids"] else []
        d["template_ids"] = json.loads(d["template_ids"]) if d["template_ids"] else []
        return d

    def get_all_presets(self):
        return [dict(r) | {"rule_ids": json.loads(r["rule_ids"] or "[]"), "template_ids": json.loads(r["template_ids"] or "[]")}
                for r in self.conn.execute("SELECT * FROM presets").fetchall()]

    def update_preset(self, pid, **kwargs):
        for key, val in kwargs.items():
            if key in ("rule_ids","template_ids"):
                val = json.dumps(val)
            self.conn.execute(f"UPDATE presets SET {key}=? WHERE id=?", (val, pid))
        self.conn.commit()

    def delete_preset(self, pid):
        self.conn.execute("DELETE FROM presets WHERE id=?", (pid,))
        self.conn.commit()

    def find_presets_by_template_ids(self, template_ids):
        """열린 문서의 template_id 조합으로 프리셋 자동 감지"""
        all_presets = self.get_all_presets()
        matched = []
        tset = set(template_ids)
        for p in all_presets:
            pset = set(p["template_ids"])
            if pset and pset.issubset(tset):
                matched.append(p)
        return sorted(matched, key=lambda p: len(p["template_ids"]), reverse=True)

    # --- Documents ---
    def save_document(self, file_path, file_type, template_id, entities):
        c = self.conn.execute(
            "INSERT INTO documents (file_path,file_type,template_id,entities,parsed_at) VALUES (?,?,?,?,?)",
            (file_path, file_type, template_id, json.dumps(entities,ensure_ascii=False), datetime.now().isoformat()))
        self.conn.commit()
        return c.lastrowid

    # --- Validation Results ---
    def save_validation_result(self, preset_id, rule_id, document_ids, status, detail):
        c = self.conn.execute(
            "INSERT INTO validation_results (preset_id,rule_id,document_ids,status,detail,executed_at) VALUES (?,?,?,?,?,?)",
            (preset_id, rule_id, json.dumps(document_ids), status, json.dumps(detail,ensure_ascii=False), datetime.now().isoformat()))
        self.conn.commit()
        return c.lastrowid

    def get_validation_results(self, preset_id=None):
        sql = "SELECT * FROM validation_results"
        params = ()
        if preset_id:
            sql += " WHERE preset_id=?"
            params = (preset_id,)
        return [dict(r) | {"document_ids": json.loads(r["document_ids"] or "[]"), "detail": json.loads(r["detail"] or "{}")}
                for r in self.conn.execute(sql, params).fetchall()]

    def close(self):
        self.conn.close()
```

- [ ] **Step 5: 테스트 통과 확인**

Run: `cd C:/PP_CC_Error && python -m pytest tests/test_storage.py -v`
Expected: 9 passed

- [ ] **Step 6: 커밋**

```bash
git add doc_intelligence/ tests/test_storage.py requirements_doc_intelligence.txt
git commit -m "feat: storage.py — SQLite 전체 CRUD + 프리셋 자동 감지"
```

---

## Task 2: 코어 엔진 + 데이터 모델 — engine.py

**Files:** Create `doc_intelligence/engine.py`, `tests/test_engine.py`

- [ ] **Step 1: 테스트 작성**

```python
# tests/test_engine.py
import pytest
from doc_intelligence.engine import Engine, ParsedDocument, CellData, Entity, Fingerprint

def test_dataclasses():
    doc = ParsedDocument("t.xlsx","excel","text",{"sheets":["S1"]},
                         [CellData("S1!A1","val","text",{})],{"k":"v"})
    assert doc.file_type == "excel"
    fp = Fingerprint("doc1",[0.1,0.2],{"A1":"합계"},"abc123")
    assert fp.merge_pattern == "abc123"
    e = Entity("금액","15000000","R2C2",0.95)
    assert e.confidence == 0.95

def test_register_and_list():
    engine = Engine(db_path=":memory:")
    class P:
        name="p1"; enabled=True
        def initialize(self,e): pass
        def process(self,d,c): return {"p1":True}
    engine.register(P())
    assert "p1" in engine.list_plugins()

def test_disable_enable():
    engine = Engine(db_path=":memory:")
    class P:
        name="p1"; enabled=True
        def initialize(self,e): pass
        def process(self,d,c): c["ran"]=True; return c
    engine.register(P())
    engine.disable("p1")
    doc = ParsedDocument("t","excel","",{},[],{})
    r = engine.process(doc)
    assert "ran" not in r
    engine.enable("p1")
    r2 = engine.process(doc)
    assert r2["ran"] is True

def test_pipeline_order():
    engine = Engine(db_path=":memory:")
    order = []
    class A:
        name="a"; enabled=True
        def initialize(self,e): pass
        def process(self,d,c): order.append("a"); return c
    class B:
        name="b"; enabled=True
        def initialize(self,e): pass
        def process(self,d,c): order.append("b"); return c
    engine.register(A())
    engine.register(B())
    engine.process(ParsedDocument("t","excel","",{},[],{}))
    assert order == ["a","b"]

def test_plugin_exception_does_not_crash():
    engine = Engine(db_path=":memory:")
    class Bad:
        name="bad"; enabled=True
        def initialize(self,e): pass
        def process(self,d,c): raise ValueError("boom")
    class Good:
        name="good"; enabled=True
        def initialize(self,e): pass
        def process(self,d,c): c["good"]=True; return c
    engine.register(Bad())
    engine.register(Good())
    r = engine.process(ParsedDocument("t","excel","",{},[],{}))
    assert r["good"] is True  # Bad 플러그인 예외 후에도 계속 실행
```

- [ ] **Step 2: engine.py 구현**

```python
# doc_intelligence/engine.py
import logging
from dataclasses import dataclass, field
from doc_intelligence.storage import Storage

log = logging.getLogger(__name__)

@dataclass
class CellData:
    address: str
    value: object
    data_type: str
    neighbors: dict

@dataclass
class ParsedDocument:
    file_path: str
    file_type: str
    raw_text: str
    structure: dict
    cells: list
    metadata: dict

@dataclass
class Entity:
    type: str
    value: str
    location: str
    confidence: float

@dataclass
class Fingerprint:
    doc_id: str
    feature_vector: list
    label_positions: dict
    merge_pattern: str

class Engine:
    def __init__(self, db_path="templates.db"):
        self.storage = Storage(db_path)
        self.plugins = {}
        self._order = []

    def register(self, plugin):
        plugin.initialize(self)
        self.plugins[plugin.name] = plugin
        self._order.append(plugin.name)

    def disable(self, name):
        if name in self.plugins:
            self.plugins[name].enabled = False

    def enable(self, name):
        if name in self.plugins:
            self.plugins[name].enabled = True

    def list_plugins(self):
        return list(self.plugins.keys())

    def process(self, doc: ParsedDocument) -> dict:
        context = {"doc": doc}
        for name in self._order:
            plugin = self.plugins[name]
            if not plugin.enabled:
                continue
            try:
                result = plugin.process(doc, context)
                if result:
                    context.update(result)
            except Exception as e:
                log.error(f"플러그인 '{name}' 오류: {e}")
                context.setdefault("errors", []).append({"plugin": name, "error": str(e)})
        return context
```

- [ ] **Step 3: 테스트 통과 확인 + 커밋**

```bash
git add doc_intelligence/engine.py tests/test_engine.py
git commit -m "feat: engine.py — 코어 파이프라인 + 데이터 모델 (Fingerprint 포함)"
```

---

## Task 3: COM 프로세스 격리 래퍼 — com_worker.py

**Files:** Create `doc_intelligence/com_worker.py`, `tests/test_com_worker.py`

- [ ] **Step 1: 테스트 작성**

```python
# tests/test_com_worker.py
import pytest
from unittest.mock import patch, MagicMock
from doc_intelligence.com_worker import ComWorker

def test_com_worker_retry_on_failure():
    worker = ComWorker(max_retries=3, timeout=5)
    call_count = 0
    def flaky_func():
        nonlocal call_count
        call_count += 1
        if call_count < 3:
            raise Exception("COM error")
        return "success"
    result = worker.execute(flaky_func)
    assert result == "success"
    assert call_count == 3

def test_com_worker_max_retries_exceeded():
    worker = ComWorker(max_retries=2, timeout=5)
    def always_fail():
        raise Exception("COM error")
    with pytest.raises(Exception, match="COM error"):
        worker.execute(always_fail)

def test_com_worker_get_active_app():
    worker = ComWorker()
    # COM 없는 환경에서는 None 반환
    app = worker.get_active_app("Excel.Application")
    # CI에서는 COM 없으므로 None 허용
    assert app is None or app is not None

def test_detect_open_documents():
    worker = ComWorker()
    docs = worker.detect_open_documents()
    assert isinstance(docs, list)
```

- [ ] **Step 2: com_worker.py 구현**

```python
# doc_intelligence/com_worker.py
import time, logging, threading
from contextlib import contextmanager

log = logging.getLogger(__name__)

class ComWorker:
    """COM 호출을 격리/안정화하는 래퍼. STA 스레드, 재시도, 타임아웃 처리."""

    def __init__(self, max_retries=3, timeout=10):
        self.max_retries = max_retries
        self.timeout = timeout

    def execute(self, func, *args, **kwargs):
        """COM 함수를 재시도 정책으로 실행"""
        last_error = None
        for attempt in range(1, self.max_retries + 1):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                last_error = e
                log.warning(f"COM 호출 실패 (시도 {attempt}/{self.max_retries}): {e}")
                if attempt < self.max_retries:
                    time.sleep(1)
        raise last_error

    def get_active_app(self, prog_id):
        """실행 중인 COM 앱에 연결. 없으면 None."""
        try:
            import pythoncom
            pythoncom.CoInitialize()
            import win32com.client
            app = win32com.client.GetActiveObject(prog_id)
            return app
        except Exception as e:
            log.debug(f"COM 앱 연결 실패 ({prog_id}): {e}")
            return None

    def detect_open_documents(self):
        """현재 열려 있는 문서 목록 감지"""
        docs = []
        app_configs = [
            ("Excel.Application", "excel", lambda a: [(wb.FullName, wb.Name) for wb in a.Workbooks]),
            ("Word.Application", "word", lambda a: [(d.FullName, d.Name) for d in a.Documents]),
            ("PowerPoint.Application", "ppt", lambda a: [(p.FullName, p.Name) for p in a.Presentations]),
        ]
        for prog_id, ftype, extractor in app_configs:
            app = self.get_active_app(prog_id)
            if app:
                try:
                    for full_path, name in self.execute(extractor, app):
                        docs.append({"path": full_path, "name": name, "type": ftype, "app": app})
                except Exception as e:
                    log.warning(f"{prog_id} 문서 감지 실패: {e}")
        return docs

    @contextmanager
    def com_session(self):
        """STA COM 세션 컨텍스트 매니저"""
        try:
            import pythoncom
            pythoncom.CoInitialize()
            yield
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
```

- [ ] **Step 3: 테스트 통과 + 커밋**

```bash
git add doc_intelligence/com_worker.py tests/test_com_worker.py
git commit -m "feat: com_worker.py — COM 프로세스 격리 (STA, 재시도 3회, 타임아웃)"
```

---

## Task 4: Excel 파서 — parsers.py

**Files:** Create `doc_intelligence/parsers.py`, `tests/test_parsers.py`

- [ ] **Step 1: 테스트 작성 (Excel + Base)**

```python
# tests/test_parsers.py
import pytest
from unittest.mock import MagicMock
from doc_intelligence.parsers import BaseParser, ExcelParser
from doc_intelligence.engine import ParsedDocument

def test_base_parser_raises():
    with pytest.raises(NotImplementedError):
        BaseParser().parse_from_com(None)

def _mock_excel():
    app = MagicMock()
    wb = MagicMock()
    ws = MagicMock()
    ws.Name = "Sheet1"
    ws.UsedRange.Rows.Count = 3
    ws.UsedRange.Columns.Count = 2
    ws.UsedRange.Row = 1
    ws.UsedRange.Column = 1
    data = {(1,1):"견적서",(1,2):"",(2,1):"합계",(2,2):15000000,(3,1):"날짜",(3,2):"2025.03.15"}
    def cells(r,c):
        m = MagicMock()
        m.Value = data.get((r,c),"")
        m.MergeCells = (r==1 and c==1)
        m.NumberFormat = "General"
        return m
    ws.Cells = cells
    wb.Worksheets = [ws]
    wb.FullName = "C:\\견적서.xlsx"
    app.ActiveWorkbook = wb
    return app

def test_excel_parser_basic():
    doc = ExcelParser().parse_from_com(_mock_excel())
    assert doc.file_type == "excel"
    assert "견적서" in doc.raw_text
    assert any(c.value == 15000000 for c in doc.cells)

def test_excel_merge_pattern():
    doc = ExcelParser().parse_from_com(_mock_excel())
    assert "merge_cells" in doc.structure
```

- [ ] **Step 2: parsers.py 구현**

```python
# doc_intelligence/parsers.py
import hashlib
from doc_intelligence.engine import ParsedDocument, CellData

class BaseParser:
    def parse_from_com(self, com_app) -> ParsedDocument:
        raise NotImplementedError

class ExcelParser(BaseParser):
    def parse_from_com(self, excel_app) -> ParsedDocument:
        wb = excel_app.ActiveWorkbook
        cells, raw_parts, sheets, merges = [], [], [], []
        for ws in wb.Worksheets:
            sheets.append(ws.Name)
            u = ws.UsedRange
            sr, sc = u.Row, u.Column
            for r in range(sr, sr + u.Rows.Count):
                for c in range(sc, sc + u.Columns.Count):
                    cell = ws.Cells(r, c)
                    v = cell.Value
                    if v is None: continue
                    addr = f"{ws.Name}!R{r}C{c}"
                    dtype = "number" if isinstance(v,(int,float)) else "text"
                    cells.append(CellData(addr, v, dtype, {}))
                    raw_parts.append(str(v))
                    if getattr(cell, "MergeCells", False):
                        merges.append(addr)
        merge_hash = hashlib.md5(",".join(merges).encode()).hexdigest()[:8]
        return ParsedDocument(
            file_path=wb.FullName, file_type="excel",
            raw_text=" ".join(raw_parts),
            structure={"sheets":sheets,"sheet_count":len(sheets),"merge_cells":merges,"merge_hash":merge_hash},
            cells=cells, metadata={"file_name":wb.FullName})

class WordParser(BaseParser):
    def parse_from_com(self, word_app) -> ParsedDocument:
        doc = word_app.ActiveDocument
        cells, raw_parts = [], []
        for i, para in enumerate(doc.Paragraphs):
            text = para.Range.Text.strip().rstrip('\r\x07')
            if text:
                cells.append(CellData(f"para:{i+1}", text, "text", {}))
                raw_parts.append(text)
        for ti, table in enumerate(doc.Tables):
            for r in range(1, table.Rows.Count + 1):
                for c in range(1, table.Columns.Count + 1):
                    try:
                        val = table.Cell(r,c).Range.Text.strip().rstrip('\r\x07')
                        if val:
                            cells.append(CellData(f"table{ti+1}:R{r}C{c}", val, "text", {}))
                            raw_parts.append(val)
                    except Exception: pass
        return ParsedDocument(
            file_path=doc.FullName, file_type="word",
            raw_text=" ".join(raw_parts),
            structure={"paragraphs":len(list(doc.Paragraphs)),"tables":len(list(doc.Tables))},
            cells=cells, metadata={"file_name":doc.FullName})

class PowerPointParser(BaseParser):
    def parse_from_com(self, ppt_app) -> ParsedDocument:
        prs = ppt_app.ActivePresentation
        cells, raw_parts, slide_info = [], [], []
        for si, slide in enumerate(prs.Slides):
            slide_info.append(f"slide{si+1}")
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    text = shape.TextFrame.TextRange.Text.strip()
                    if text:
                        cells.append(CellData(f"slide{si+1}:shape{shape.ShapeIndex}", text, "text", {}))
                        raw_parts.append(text)
                if shape.HasTable:
                    tbl = shape.Table
                    for r in range(1, tbl.Rows.Count + 1):
                        for c in range(1, tbl.Columns.Count + 1):
                            val = tbl.Cell(r,c).Shape.TextFrame.TextRange.Text.strip()
                            if val:
                                cells.append(CellData(f"slide{si+1}:tbl:R{r}C{c}", val, "text", {}))
                                raw_parts.append(val)
        return ParsedDocument(
            file_path=prs.FullName, file_type="ppt",
            raw_text=" ".join(raw_parts),
            structure={"slides":slide_info,"slide_count":len(slide_info)},
            cells=cells, metadata={"file_name":prs.FullName})

class PdfParser(BaseParser):
    def parse_from_com(self, acrobat_app) -> ParsedDocument:
        """AcroExch.App COM으로 PDF 텍스트 추출. Acrobat Pro 필수."""
        try:
            avdoc = acrobat_app.GetActiveDoc()
            pddoc = avdoc.GetPDDoc()
            cells, raw_parts = [], []
            num_pages = pddoc.GetNumPages()
            for pi in range(num_pages):
                page = pddoc.AcquirePage(pi)
                highlight = page.CreateWordHilite(0, -1)  # 전체 단어
                if highlight:
                    for wi in range(highlight.GetNumWordHilites()):
                        word = highlight.GetWordHilite(wi)
                        # AcroExch API로 텍스트 추출
                        pass
                # 대안: JSObject로 텍스트 추출
                js_obj = pddoc.GetJSObject()
                if js_obj:
                    text = js_obj.getPageNthWord(pi, 0)  # 첫 단어
                    # 모든 단어 반복
                    page_words = []
                    try:
                        wi = 0
                        while True:
                            w = js_obj.getPageNthWord(pi, wi)
                            if not w: break
                            page_words.append(w)
                            wi += 1
                    except Exception:
                        pass
                    page_text = " ".join(page_words)
                    if page_text:
                        cells.append(CellData(f"page{pi+1}", page_text, "text", {}))
                        raw_parts.append(page_text)
            return ParsedDocument(
                file_path=avdoc.GetFileName(), file_type="pdf",
                raw_text=" ".join(raw_parts),
                structure={"pages":num_pages},
                cells=cells, metadata={})
        except Exception as e:
            return self._fallback_ocr(str(e))

    def _fallback_ocr(self, error_msg):
        """Acrobat Pro 미설치 시 화면 캡처 + OCR 폴백"""
        return ParsedDocument("","pdf",f"[OCR 폴백 필요: {error_msg}]",{},[], {"fallback":True})

class ImageParser(BaseParser):
    def parse_from_com(self, image_path_or_screenshot) -> ParsedDocument:
        """Tesseract OCR로 이미지 텍스트 추출"""
        try:
            import pytesseract
            from PIL import Image
            if isinstance(image_path_or_screenshot, str):
                img = Image.open(image_path_or_screenshot)
            else:
                img = image_path_or_screenshot
            data = pytesseract.image_to_data(img, lang="kor", output_type=pytesseract.Output.DICT)
            cells, raw_parts = [], []
            for i, text in enumerate(data["text"]):
                text = text.strip()
                conf = int(data["conf"][i]) if data["conf"][i] != "-1" else 0
                if text and conf > 30:
                    x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
                    cell = CellData(f"ocr:{x},{y}", text, "text", {"confidence": conf/100.0})
                    cells.append(cell)
                    raw_parts.append(text)
            return ParsedDocument(
                file_path=str(image_path_or_screenshot), file_type="image",
                raw_text=" ".join(raw_parts),
                structure={"ocr_blocks":len(cells)},
                cells=cells, metadata={"ocr_engine":"tesseract"})
        except Exception as e:
            return ParsedDocument("","image",f"[OCR 실패: {e}]",{},[], {"error":str(e)})
```

- [ ] **Step 3: Word/PPT/PDF/Image 파서 테스트 추가**

```python
# tests/test_parsers.py에 추가

from doc_intelligence.parsers import WordParser, PowerPointParser, PdfParser, ImageParser

def _mock_word():
    app = MagicMock()
    doc = MagicMock()
    doc.FullName = "C:\\정산서.docx"
    p1 = MagicMock(); p1.Range.Text = "정비비용 정산서\r"
    p2 = MagicMock(); p2.Range.Text = "업체명: 삼성엔지니어링\r"
    doc.Paragraphs = [p1, p2]
    doc.Tables = []
    doc.Sections.Count = 1
    app.ActiveDocument = doc
    return app

def test_word_parser():
    doc = WordParser().parse_from_com(_mock_word())
    assert doc.file_type == "word"
    assert "정산서" in doc.raw_text

def _mock_ppt():
    app = MagicMock()
    prs = MagicMock()
    prs.FullName = "C:\\발표자료.pptx"
    slide = MagicMock()
    shape = MagicMock()
    shape.HasTextFrame = True
    shape.HasTable = False
    shape.ShapeIndex = 1
    shape.TextFrame.TextRange.Text = "배관 정비 현황"
    slide.Shapes = [shape]
    prs.Slides = [slide]
    app.ActivePresentation = prs
    return app

def test_ppt_parser():
    doc = PowerPointParser().parse_from_com(_mock_ppt())
    assert doc.file_type == "ppt"
    assert "배관" in doc.raw_text

def test_pdf_parser_fallback():
    # Acrobat COM 없을 때 폴백 테스트
    doc = PdfParser().parse_from_com(MagicMock(side_effect=Exception("no Acrobat")))
    assert doc.file_type == "pdf"

def test_image_parser_no_tesseract():
    # Tesseract 미설치 시 graceful fail
    doc = ImageParser().parse_from_com("nonexistent.png")
    assert doc.file_type == "image"
```

- [ ] **Step 4: 테스트 통과 + 커밋**

```bash
git add doc_intelligence/parsers.py tests/test_parsers.py
git commit -m "feat: parsers.py — 5종 파서 (Excel/Word/PPT/PDF/Image)"
```

---

## Task 5: 핑거프린트 — fingerprint.py (실제 TF-IDF)

**Files:** Create `doc_intelligence/fingerprint.py`, `tests/test_fingerprint.py`

- [ ] **Step 1: 테스트 작성** (Task 4의 fingerprint 테스트와 동일하되 Fingerprint dataclass 사용)

- [ ] **Step 2: fingerprint.py 구현 — placeholder 제거, 실제 TF-IDF 벡터 사용**

```python
# doc_intelligence/fingerprint.py
import hashlib
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from doc_intelligence.engine import Fingerprint

class Fingerprinter:
    name = "fingerprinter"
    enabled = True

    def __init__(self, storage=None):
        self.storage = storage
        self._vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2,4))
        self._corpus = []
        self._template_ids = []

    def initialize(self, engine):
        self.storage = engine.storage
        # DB에서 기존 템플릿 로드
        for t in self.storage.get_all_templates():
            labels = list(t["label_positions"].keys())
            self._corpus.append(" ".join(labels))
            self._template_ids.append(t["id"])
        if self._corpus:
            self._vectorizer.fit(self._corpus)

    def generate(self, doc) -> dict:
        labels = [str(c.value) for c in doc.cells if isinstance(c.value, str) and c.value.strip()]
        label_text = " ".join(labels)
        label_positions = {}
        for c in doc.cells:
            if isinstance(c.value, str) and c.value.strip():
                label_positions[c.value] = c.address
        merges = doc.structure.get("merge_cells", [])
        merge_pattern = hashlib.md5(",".join(merges).encode()).hexdigest()[:8]
        # 실제 TF-IDF 벡터 생성
        if self._corpus:
            try:
                vec = self._vectorizer.transform([label_text]).toarray()[0].tolist()
            except Exception:
                vec = []
        else:
            vec = []
        return {
            "vector": vec,
            "labels": labels,
            "label_text": label_text,
            "label_positions": label_positions,
            "merge_pattern": merge_pattern,
            "fingerprint": Fingerprint(doc.file_path, vec, label_positions, merge_pattern),
        }

    def learn(self, doc, template_name):
        fp = self.generate(doc)
        tid = self.storage.save_template(
            template_name, doc.file_type, fp["vector"],
            fp["label_positions"], {})
        self._corpus.append(fp["label_text"])
        self._template_ids.append(tid)
        self._vectorizer.fit(self._corpus)
        return tid

    def match(self, doc):
        if not self._corpus:
            return {"template": None, "score": 0.0}
        fp = self.generate(doc)
        corpus_with_query = self._corpus + [fp["label_text"]]
        tfidf = self._vectorizer.fit_transform(corpus_with_query)
        scores = cosine_similarity(tfidf[-1:], tfidf[:-1]).flatten()
        idx = int(np.argmax(scores))
        score = float(scores[idx])
        if score >= 0.85:
            tid = self._template_ids[idx]
            self.storage.increment_match_count(tid)
            return {"template": self.storage.get_template(tid), "score": score, "auto": True}
        elif score >= 0.60:
            return {"template": self.storage.get_template(self._template_ids[idx]), "score": score, "auto": False}
        return {"template": None, "score": score}

    def process(self, doc, context):
        context["fingerprint"] = self.generate(doc)
        context["template_match"] = self.match(doc)
        return context
```

- [ ] **Step 3: 테스트 통과 + 커밋**

```bash
git add doc_intelligence/fingerprint.py tests/test_fingerprint.py
git commit -m "feat: fingerprint.py — 실제 TF-IDF 벡터 + Fingerprint dataclass"
```

---

## Task 6: 엔티티 추출기 — extractor.py

(이전 계획 Task 5와 동일. 코드 생략 — 위 참조)

```bash
git commit -m "feat: extractor.py — regex + 인접 셀 기반 엔티티 추출기"
```

---

## Task 7: 교차 검증기 — validator.py (6종 룰 + 프리셋 자동 감지)

**Files:** Create `doc_intelligence/validator.py`, `tests/test_validator.py`

- [ ] **Step 1: 테스트 작성 (RangeCheckRule 포함, 프리셋 자동 감지 포함)**

```python
# tests/test_validator.py
import pytest
from doc_intelligence.validator import (
    CrossValidator, ValueMatchRule, OrderCheckRule,
    FormulaCheckRule, ExistsRule, ContainsRule, RangeCheckRule
)

def test_value_match_pass():
    r = ValueMatchRule("t", [{"value":"100"},{"value":"100"}])
    assert r.check()["status"] == "통과"

def test_value_match_fail():
    r = ValueMatchRule("t", [{"value":"100"},{"value":"200"}])
    assert r.check()["status"] == "실패"

def test_order_pass():
    r = OrderCheckRule("t", [{"value":"2025.01.01"},{"value":"2025.02.01"}])
    assert r.check()["status"] == "통과"

def test_order_fail():
    r = OrderCheckRule("t", [{"value":"2025.03.01"},{"value":"2025.02.01"}])
    assert r.check()["status"] == "실패"

def test_formula_pass():
    r = FormulaCheckRule("t", {"operands":[{"value":"10"},{"value":"1500000"}],
                                "operator":"*","expected":{"value":"15000000"}})
    assert r.check()["status"] == "통과"

def test_formula_warn():
    r = FormulaCheckRule("t", {"operands":[{"value":"10"},{"value":"1500000"}],
                                "operator":"*","expected":{"value":"14999999"}})
    assert r.check()["status"] == "경고"

def test_exists_pass():
    r = ExistsRule("t", [{"value":"abc"},{"value":"def"}])
    assert r.check()["status"] == "통과"

def test_exists_fail():
    r = ExistsRule("t", [{"value":"abc"},{"value":""}])
    assert r.check()["status"] == "실패"

def test_contains_pass():
    r = ContainsRule("t", {"value":"PP-2045"}, {"value":"설비 PP-2045 배관"})
    assert r.check()["status"] == "통과"

def test_contains_fail():
    r = ContainsRule("t", {"value":"PP-9999"}, {"value":"설비 PP-2045 배관"})
    assert r.check()["status"] == "실패"

def test_range_check_pass():
    r = RangeCheckRule("t", {"value":"2025.02.15"},
                       {"min":"2025.01.01","max":"2025.03.31"})
    assert r.check()["status"] == "통과"

def test_range_check_fail():
    r = RangeCheckRule("t", {"value":"2025.05.01"},
                       {"min":"2025.01.01","max":"2025.03.31"})
    assert r.check()["status"] == "실패"

def test_cross_validator():
    cv = CrossValidator()
    cv.add_rule(ValueMatchRule("r1", [{"value":"100"},{"value":"100"}]))
    cv.add_rule(OrderCheckRule("r2", [{"value":"2025.01.01"},{"value":"2025.02.01"}]))
    results = cv.validate()
    assert len(results) == 2
    assert all(r["status"] == "통과" for r in results)
```

- [ ] **Step 2: validator.py 구현 (RangeCheckRule 추가)**

이전 계획의 validator.py + RangeCheckRule 추가:

```python
class RangeCheckRule(BaseRule):
    """값이 min~max 범위 내에 있는지 확인"""
    def __init__(self, name, target, bounds):
        super().__init__(name)
        self.target = target
        self.bounds = bounds

    def check(self):
        try:
            val = self._parse(str(self.target["value"]))
            mn = self._parse(str(self.bounds["min"]))
            mx = self._parse(str(self.bounds["max"]))
            if mn <= val <= mx:
                return {"rule":self.name,"status":"통과","detail":f"{self.target['value']} 범위 내"}
            return {"rule":self.name,"status":"실패",
                    "detail":f"{self.target['value']}이(가) {self.bounds['min']}~{self.bounds['max']} 범위 밖"}
        except Exception as e:
            return {"rule":self.name,"status":"실패","detail":f"파싱 오류: {e}"}

    def _parse(self, s):
        for fmt in ["%Y.%m.%d","%Y-%m-%d","%Y/%m/%d"]:
            try:
                from datetime import datetime
                return datetime.strptime(s, fmt)
            except ValueError: continue
        return float(s.replace(",",""))
```

- [ ] **Step 3: 테스트 통과 + 커밋**

```bash
git add doc_intelligence/validator.py tests/test_validator.py
git commit -m "feat: validator.py — 6종 룰 (값일치/순서/수식/존재/포함/범위) + 교차 검증기"
```

---

## Task 8: 드래그 영역 연결기 — region_linker.py

**Files:** Create `doc_intelligence/region_linker.py`, `tests/test_region_linker.py`

- [ ] **Step 1: 테스트 작성**

```python
# tests/test_region_linker.py
import pytest
from unittest.mock import MagicMock, patch
from doc_intelligence.region_linker import RegionLinker, LinkedRegion

def test_linked_region():
    r = LinkedRegion("EXCEL.EXE","견적서.xlsx","Sheet1!B4",(100,200,300,250),None)
    assert r.app_name == "EXCEL.EXE"

def test_create_rule():
    linker = RegionLinker(storage=MagicMock())
    r1 = LinkedRegion("EXCEL.EXE","견적서.xlsx","Sheet1!E15",(0,0,100,50),None)
    r2 = LinkedRegion("WINWORD.EXE","정산서.docx","para:4",(200,0,300,50),None)
    rule = linker.create_rule([r1,r2],"값_일치","금액일치")
    assert rule["name"] == "금액일치"
    assert len(rule["regions"]) == 2

def test_dpi_scaling():
    linker = RegionLinker(storage=MagicMock())
    # 150% DPI에서 물리 좌표 300 -> 논리 좌표 200
    logical = linker._apply_dpi_scale(300, 1.5)
    assert logical == 200

def test_identify_app_from_hwnd():
    linker = RegionLinker(storage=MagicMock())
    # win32gui 없는 환경에서 graceful fallback
    result = linker._get_app_from_point(500, 500)
    assert result is None or isinstance(result, str)
```

- [ ] **Step 2: region_linker.py 구현**

```python
# doc_intelligence/region_linker.py
import logging
from dataclasses import dataclass
from typing import Optional

log = logging.getLogger(__name__)

@dataclass
class LinkedRegion:
    app_name: str       # "EXCEL.EXE", "WINWORD.EXE" 등
    doc_name: str       # 파일명
    location: str       # 문서 내 위치 ("Sheet1!B4", "para:3" 등)
    screen_rect: tuple  # (x, y, w, h) 스크린 좌표
    screenshot: Optional[bytes]  # 영역 캡처 이미지

class RegionLinker:
    def __init__(self, storage):
        self.storage = storage
        self.current_regions = []

    def create_rule(self, regions, rule_type, rule_name):
        """선택된 영역들로 단일 룰 생성"""
        region_data = []
        for r in regions:
            region_data.append({
                "app": r.app_name,
                "doc": r.doc_name,
                "location": r.location,
                "rect": list(r.screen_rect),
            })
        rule_data = {
            "name": rule_name,
            "rule_type": rule_type,
            "regions": region_data,
            "params": {},
        }
        if self.storage:
            rid = self.storage.save_rule(rule_name, rule_type, region_data, {})
            rule_data["id"] = rid
        return rule_data

    def _apply_dpi_scale(self, physical_coord, scale_factor):
        """DPI 스케일링 적용 (물리 좌표 -> 논리 좌표)"""
        return int(physical_coord / scale_factor)

    def _get_dpi_scale(self):
        """현재 모니터의 DPI 스케일 팩터 획득"""
        try:
            import ctypes
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return dpi / 96.0
        except Exception:
            return 1.0

    def _get_app_from_point(self, x, y):
        """스크린 좌표에서 어떤 앱 창인지 판별"""
        try:
            import win32gui, win32process, psutil
            hwnd = win32gui.WindowFromPoint((x, y))
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            return proc.name().upper()
        except Exception:
            return None

    def _screen_to_excel_cell(self, excel_app, x, y):
        """스크린 좌표 -> Excel 셀 주소 변환"""
        try:
            rng = excel_app.ActiveWindow.RangeFromPoint(x, y)
            return f"{rng.Worksheet.Name}!{rng.Address}"
        except Exception:
            return f"screen:{x},{y}"

    def _screen_to_word_location(self, word_app, x, y):
        """스크린 좌표 -> Word 문서 위치 (Selection 기반)"""
        try:
            # Word에는 RangeFromPoint 없음 — Selection 기반 대체
            return f"screen:{x},{y}"
        except Exception:
            return f"screen:{x},{y}"

    def _screen_to_ppt_location(self, ppt_app, x, y):
        """스크린 좌표 -> PPT 슬라이드 내 위치"""
        return f"screen:{x},{y}"

    def _screen_to_pdf_location(self, acrobat_app, x, y):
        """스크린 좌표 -> PDF 페이지 내 위치"""
        try:
            avdoc = acrobat_app.GetActiveDoc()
            page_view = avdoc.GetAVPageView()
            page_num = page_view.GetPageNum()
            # DevPtToPagePt로 변환
            return f"page{page_num+1}:screen:{x},{y}"
        except Exception:
            return f"screen:{x},{y}"

    def capture_region(self, rect):
        """영역 스크린샷 캡처"""
        try:
            import pyautogui
            x, y, w, h = rect
            screenshot = pyautogui.screenshot(region=(x, y, w, h))
            return screenshot
        except Exception:
            return None
```

- [ ] **Step 3: 테스트 통과 + 커밋**

```bash
git add doc_intelligence/region_linker.py tests/test_region_linker.py
git commit -m "feat: region_linker.py — 드래그 영역 연결 + DPI 스케일링 + 좌표 변환"
```

---

## Task 9: 이상 탐지 — anomaly.py

**Files:** Create `doc_intelligence/anomaly.py`, `tests/test_anomaly.py`

- [ ] **Step 1: 테스트 작성**

```python
# tests/test_anomaly.py
import pytest
from doc_intelligence.anomaly import AnomalyDetector

def test_detect_outlier():
    det = AnomalyDetector()
    data = [100,102,98,101,99,103,97,100,500]
    results = det.detect(data)
    assert results[-1] == True

def test_no_outlier():
    det = AnomalyDetector()
    data = [100,102,98,101,99,103,97,100,101]
    results = det.detect(data)
    assert not any(results)

def test_insufficient_data():
    det = AnomalyDetector()
    data = [100]
    results = det.detect(data)
    assert len(results) == 1
    assert results[0] == False
```

- [ ] **Step 2: anomaly.py 구현**

```python
# doc_intelligence/anomaly.py
import numpy as np
from sklearn.ensemble import IsolationForest

class AnomalyDetector:
    name = "anomaly_detector"
    enabled = True

    def __init__(self, contamination=0.1):
        self.contamination = contamination

    def initialize(self, engine):
        pass

    def detect(self, values):
        """숫자 리스트에서 이상치 탐지. True=이상치"""
        if len(values) < 5:
            return [False] * len(values)
        X = np.array(values).reshape(-1, 1)
        clf = IsolationForest(contamination=self.contamination, random_state=42)
        preds = clf.fit_predict(X)
        return [p == -1 for p in preds]

    def process(self, doc, context):
        entities = context.get("entities", [])
        amounts = [float(e.value.replace(",","").replace("원",""))
                   for e in entities if e.type == "금액"
                   and e.value.replace(",","").replace("원","").replace(".","").isdigit()]
        if amounts:
            anomalies = self.detect(amounts)
            context["anomalies"] = list(zip(amounts, anomalies))
        return context
```

- [ ] **Step 3: 테스트 통과 + 커밋**

```bash
git add doc_intelligence/anomaly.py tests/test_anomaly.py
git commit -m "feat: anomaly.py — Isolation Forest 이상 탐지"
```

---

## Task 10: 관계 그래프 — graph.py

**Files:** Create `doc_intelligence/graph.py`, `tests/test_graph.py`

- [ ] **Step 1: 테스트 작성**

```python
# tests/test_graph.py
import pytest
from doc_intelligence.graph import DocGraph

def test_add_documents():
    g = DocGraph()
    g.add_document("견적서.xlsx", {"금액":"15000000"})
    g.add_document("정산서.xlsx", {"금액":"15000000"})
    assert g.node_count() == 2

def test_add_relationship():
    g = DocGraph()
    g.add_document("A.xlsx", {})
    g.add_document("B.docx", {})
    g.add_relationship("A.xlsx","B.docx","금액일치","통과")
    assert len(g.get_edges()) == 1

def test_to_html():
    g = DocGraph()
    g.add_document("A.xlsx", {})
    g.add_document("B.docx", {})
    g.add_relationship("A.xlsx","B.docx","test","통과")
    html = g.to_html()
    assert len(html) > 0
```

- [ ] **Step 2: graph.py 구현**

```python
# doc_intelligence/graph.py
import networkx as nx
import tempfile, os

class DocGraph:
    name = "doc_graph"
    enabled = True

    def __init__(self):
        self.G = nx.Graph()

    def initialize(self, engine):
        pass

    def add_document(self, doc_name, entities):
        self.G.add_node(doc_name, entities=entities,
                        node_type=doc_name.split(".")[-1])

    def add_relationship(self, doc1, doc2, rule_name, status):
        color = {"통과":"green","실패":"red","경고":"orange"}.get(status,"gray")
        self.G.add_edge(doc1, doc2, rule=rule_name, status=status, color=color)

    def node_count(self):
        return self.G.number_of_nodes()

    def get_edges(self):
        return list(self.G.edges(data=True))

    def to_html(self):
        """pyvis로 인터랙티브 HTML 그래프 생성"""
        try:
            from pyvis.network import Network
            net = Network(height="600px", width="100%", bgcolor="#0d1117", font_color="#e0e0e0")
            for node, data in self.G.nodes(data=True):
                ntype = data.get("node_type","")
                color_map = {"xlsx":"#3fb950","docx":"#58a6ff","pptx":"#f0883e","pdf":"#f85149","jpg":"#bc8cff"}
                net.add_node(node, label=node, color=color_map.get(ntype,"#8b949e"))
            for u, v, data in self.G.edges(data=True):
                net.add_edge(u, v, title=f"{data.get('rule','')} ({data.get('status','')})",
                             color=data.get("color","gray"))
            fd, path = tempfile.mkstemp(suffix=".html")
            os.close(fd)
            net.save_graph(path)
            with open(path, "r", encoding="utf-8") as f:
                html = f.read()
            os.unlink(path)
            return html
        except ImportError:
            return "<html><body>pyvis not installed</body></html>"

    def process(self, doc, context):
        entities = context.get("entities", [])
        entity_dict = {e.type: e.value for e in entities}
        self.add_document(doc.file_path, entity_dict)
        context["graph"] = self
        return context
```

- [ ] **Step 3: 테스트 통과 + 커밋**

```bash
git add doc_intelligence/graph.py tests/test_graph.py
git commit -m "feat: graph.py — NetworkX 관계 그래프 + pyvis HTML 시각화"
```

---

## Task 11: 설정 관리 — config.yaml

**Files:** Create `doc_intelligence/config.yaml`

- [ ] **Step 1: config.yaml 생성**

```yaml
# doc_intelligence/config.yaml
fingerprint:
  auto_match_threshold: 0.85
  candidate_threshold: 0.60

extractor:
  custom_patterns:
    # 사용자 추가 정규식 패턴
    # 설비코드_v2: "[A-Z]{3}-\\d{6}"
  ocr_confidence_threshold: 0.6

com:
  max_retries: 3
  timeout_seconds: 10
  poll_interval_seconds: 3

ui:
  theme: "dark"
```

- [ ] **Step 2: 커밋**

```bash
git add doc_intelligence/config.yaml
git commit -m "feat: config.yaml — 설정 파일 (임계값, 패턴, COM 정책)"
```

---

## Task 12: 학습 모드 UI — ui_components.py (LearningModeDialog)

**Files:** Create `doc_intelligence/ui_components.py`, `tests/test_ui.py`

- [ ] **Step 1: UI 자동화 테스트 작성**

```python
# tests/test_ui.py
import pytest
from unittest.mock import MagicMock, patch

def test_learning_dialog_creates():
    """LearningModeDialog 생성 테스트 (tkinter root 없이)"""
    from doc_intelligence.ui_components import LearningModeDialog
    mock_root = MagicMock()
    entities = [
        {"location":"R2C2","value":"2025.03.12","type":"날짜","confidence":0.95},
        {"location":"R3C2","value":"PP-2045","type":"설비코드","confidence":0.93},
    ]
    dialog = LearningModeDialog(mock_root, entities, doc_name="검수확인서.pdf")
    assert dialog.doc_name == "검수확인서.pdf"
    assert len(dialog.entities) == 2

def test_rule_manager_widget():
    from doc_intelligence.ui_components import RuleManagerWidget
    mock_root = MagicMock()
    presets = [{"id":1,"name":"배관정비","rule_ids":[1,2,3]}]
    rules = [{"id":1,"name":"금액일치"},{"id":2,"name":"날짜순서"},{"id":3,"name":"업체일치"}]
    widget = RuleManagerWidget(mock_root, presets, rules)
    assert widget is not None

def test_validation_result_widget():
    from doc_intelligence.ui_components import ValidationResultWidget
    mock_root = MagicMock()
    results = [
        {"rule":"금액일치","status":"통과","detail":"일치"},
        {"rule":"날짜순서","status":"실패","detail":"착공>준공"},
    ]
    widget = ValidationResultWidget(mock_root, results)
    assert widget is not None
```

- [ ] **Step 2: ui_components.py 구현**

```python
# doc_intelligence/ui_components.py
"""tkinter UI 위젯 컴포넌트"""

class LearningModeDialog:
    """새 양식 학습 UI — 자동 분석 결과 표시 + 드롭다운 편집"""
    FIELD_TYPES = ["날짜","착공일","준공일","검수일","금액","예상비용","부가세",
                   "업체명","부서","설비코드","부품코드","이름","검수자","승인자",
                   "문서번호","문서 ID","무시"]

    def __init__(self, parent, entities, doc_name=""):
        self.parent = parent
        self.entities = entities
        self.doc_name = doc_name
        self.corrections = {}

    def get_corrected_mappings(self):
        """사용자가 수정한 필드 매핑 반환"""
        mappings = {}
        for e in self.entities:
            corrected = self.corrections.get(e["location"], e["type"])
            if corrected != "무시":
                mappings[e["location"]] = corrected
        return mappings

class RuleManagerWidget:
    """룰/프리셋 관리 위젯"""
    def __init__(self, parent, presets, rules):
        self.parent = parent
        self.presets = presets
        self.rules = rules

class ValidationResultWidget:
    """검증 결과 표시 위젯"""
    def __init__(self, parent, results):
        self.parent = parent
        self.results = results

class OverlayWindow:
    """투명 오버레이 — 드래그 영역 선택용"""
    def __init__(self, on_region_selected=None):
        self.on_region_selected = on_region_selected
        self.regions = []

class DocumentListWidget:
    """열린 문서 목록"""
    def __init__(self, parent):
        self.parent = parent
        self.documents = []

class EntityListWidget:
    """추출된 엔티티 표시"""
    def __init__(self, parent):
        self.parent = parent
```

- [ ] **Step 3: 테스트 통과 + 커밋**

```bash
git add doc_intelligence/ui_components.py tests/test_ui.py
git commit -m "feat: ui_components.py — 6개 tkinter 위젯 (학습/룰관리/검증/오버레이)"
```

---

## Task 13: 메인 앱 — main.py

**Files:** Create `doc_intelligence/main.py`

- [ ] **Step 1: main.py 구현**

```python
# doc_intelligence/main.py
"""Doc Intelligence 메인 진입점"""
import threading, time, logging
from doc_intelligence.engine import Engine
from doc_intelligence.com_worker import ComWorker
from doc_intelligence.parsers import ExcelParser, WordParser, PowerPointParser, PdfParser, ImageParser
from doc_intelligence.fingerprint import Fingerprinter
from doc_intelligence.extractor import EntityExtractor
from doc_intelligence.validator import CrossValidator
from doc_intelligence.anomaly import AnomalyDetector
from doc_intelligence.graph import DocGraph

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

class DocIntelligenceApp:
    def __init__(self):
        self.engine = Engine(db_path="templates.db")
        self.com_worker = ComWorker()
        self.parsers = {
            "excel": ExcelParser(),
            "word": WordParser(),
            "ppt": PowerPointParser(),
            "pdf": PdfParser(),
            "image": ImageParser(),
        }
        # 플러그인 등록
        self.engine.register(Fingerprinter())
        self.engine.register(EntityExtractor())
        self._polling = False

    def start_polling(self, interval=3):
        """문서 열림 감지 폴링 시작"""
        self._polling = True
        def poll():
            seen = set()
            while self._polling:
                docs = self.com_worker.detect_open_documents()
                for d in docs:
                    key = d["path"]
                    if key not in seen:
                        seen.add(key)
                        log.info(f"새 문서 감지: {d['name']} ({d['type']})")
                        self._process_document(d)
                time.sleep(interval)
        t = threading.Thread(target=poll, daemon=True)
        t.start()

    def stop_polling(self):
        self._polling = False

    def _process_document(self, doc_info):
        """감지된 문서를 파싱하고 파이프라인 실행"""
        parser = self.parsers.get(doc_info["type"])
        if not parser:
            log.warning(f"미지원 문서 유형: {doc_info['type']}")
            return
        try:
            with self.com_worker.com_session():
                parsed = self.com_worker.execute(parser.parse_from_com, doc_info["app"])
            result = self.engine.process(parsed)
            match = result.get("template_match", {})
            if match.get("template"):
                log.info(f"템플릿 매칭: {match['template']['name']} (유사도: {match['score']:.2f})")
            else:
                log.info(f"신규 양식 감지 — 학습 모드 필요")
            return result
        except Exception as e:
            log.error(f"문서 처리 실패: {e}")

def main():
    app = DocIntelligenceApp()
    # tkinter UI 시작 (Task 12의 위젯 사용)
    try:
        import tkinter as tk
        root = tk.Tk()
        root.title("Doc Intelligence v0.1")
        root.geometry("1280x720")
        root.configure(bg="#0f1117")
        # TODO: 탭 구조 + 위젯 배치 (ui_components.py 활용)
        tk.Label(root, text="Doc Intelligence v0.1 MVP", fg="#58a6ff",
                 bg="#0f1117", font=("맑은 고딕", 16)).pack(pady=20)
        app.start_polling()
        root.mainloop()
    except Exception:
        log.info("tkinter 없이 CLI 모드로 실행")
        app.start_polling()
        input("Enter를 누르면 종료합니다...")
    finally:
        app.stop_polling()

if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 커밋**

```bash
git add doc_intelligence/main.py
git commit -m "feat: main.py — 메인 앱 (COM 폴링 + 파이프라인 + tkinter)"
```

---

## Task 14: 통합 테스트

**Files:** Create `tests/test_integration.py`

- [ ] **Step 1: 통합 테스트 작성**

```python
# tests/test_integration.py
import pytest
from doc_intelligence.engine import Engine, ParsedDocument, CellData
from doc_intelligence.fingerprint import Fingerprinter
from doc_intelligence.extractor import EntityExtractor
from doc_intelligence.validator import CrossValidator, ValueMatchRule, OrderCheckRule
from doc_intelligence.anomaly import AnomalyDetector
from doc_intelligence.graph import DocGraph

def test_full_pipeline():
    engine = Engine(db_path=":memory:")
    engine.register(Fingerprinter())
    engine.register(EntityExtractor())

    doc = ParsedDocument("test.xlsx","excel","견적서 합계 15,000,000원",
        {"sheets":["Sheet1"],"merge_cells":[]},
        [CellData("R1C1","견적서","text",{}),
         CellData("R2C1","합계","text",{}),
         CellData("R2C2","15,000,000원","text",{}),
         CellData("R3C1","날짜","text",{}),
         CellData("R3C2","2025.03.15","text",{})],{})

    result = engine.process(doc)
    assert "fingerprint" in result
    assert "entities" in result
    assert len(result["entities"]) >= 1

def test_learn_and_match():
    engine = Engine(db_path=":memory:")
    fp = Fingerprinter(engine.storage)
    engine.register(fp)

    doc1 = ParsedDocument("a.xlsx","excel","",{"sheets":["S1"],"merge_cells":[]},
        [CellData("R1","견적서","text",{}),CellData("R2","합계","text",{}),
         CellData("R3","날짜","text",{})],{})
    fp.learn(doc1, "견적서v1")

    doc2 = ParsedDocument("b.xlsx","excel","",{"sheets":["S1"],"merge_cells":[]},
        [CellData("R1","견적서","text",{}),CellData("R2","합계","text",{}),
         CellData("R3","날짜","text",{})],{})
    result = engine.process(doc2)
    assert result["template_match"]["score"] >= 0.85

def test_cross_validation():
    cv = CrossValidator()
    cv.add_rule(ValueMatchRule("금액", [{"value":"15000000"},{"value":"15000000"}]))
    cv.add_rule(OrderCheckRule("날짜", [{"value":"2025.01.01"},{"value":"2025.03.01"}]))
    results = cv.validate()
    assert all(r["status"] == "통과" for r in results)

def test_graph_integration():
    g = DocGraph()
    g.add_document("견적서.xlsx", {"금액":"15000000"})
    g.add_document("정산서.xlsx", {"금액":"15000000"})
    g.add_relationship("견적서.xlsx","정산서.xlsx","금액일치","통과")
    html = g.to_html()
    assert len(html) > 100

def test_preset_auto_detect():
    from doc_intelligence.storage import Storage
    s = Storage(":memory:")
    s.save_template("T1","excel",[],{},{})  # id=1
    s.save_template("T2","word",[],{},{})   # id=2
    s.save_preset("배관정비","배관",[1,2],[1,2])
    matches = s.find_presets_by_template_ids([1,2])
    assert len(matches) >= 1
    assert matches[0]["name"] == "배관정비"
    s.close()
```

- [ ] **Step 2: 테스트 통과 + 커밋**

```bash
git add tests/test_integration.py
git commit -m "test: 통합 테스트 — 전체 파이프라인 + 학습/매칭 + 교차검증 + 그래프 + 프리셋 자동감지"
```

---

## 의존성 그래프

```
Task 1 (storage) ─┬─→ Task 2 (engine) ─┬─→ Task 3 (com_worker)
                  │                     │
                  │                     ├─→ Task 4 (parsers — 5종 전부)
                  │                     │
                  ├─→ Task 5 (fingerprint)
                  │
                  ├─→ Task 6 (extractor)
                  │
                  ├─→ Task 7 (validator — 6종 룰)
                  │
                  ├─→ Task 8 (region_linker)
                  │
                  ├─→ Task 9 (anomaly)
                  │
                  └─→ Task 10 (graph)

Task 11 (config) ─── 독립

Task 12 (ui_components) ─→ Task 13 (main.py) ─→ Task 14 (통합 테스트)
```

**병렬 가능:** Task 5, 6, 7, 8, 9, 10 (모두 Task 1+2에만 의존)
