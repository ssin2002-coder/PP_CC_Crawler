# 설비일보 Word 크롤러 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 기존 word_crawler.py를 설계 문서(docs/superpowers/specs/2026-05-04-word-crawler-design.md) 기반으로 완전 재작성. 동적 표 파싱 + 항목 분리 + 저장 모드 전환 + 행 삭제 + CSV 내보내기 구현.

**Architecture:** 단일 파일(word_crawler.py) 내 섹션별 구분. DB → Parser → WordWatcher → UI(tkinter) → Tray(pystray) → Main 순서. win32com으로 열린 Word 폴링, 가장 큰 표를 동적 파싱, 2+ 개행으로 항목 분리, SQLite 비정규화 단일 테이블 저장.

**Tech Stack:** Python, win32com (pywin32), pystray, Pillow, tkinter, sqlite3

**Spec:** `docs/superpowers/specs/2026-05-04-word-crawler-design.md`

---

## 파일 구조

단일 파일 `word_crawler.py` 내 섹션:

```
word_crawler.py (전체 재작성)
├── Constants (경로, 날짜 정규식)
├── DB (init_db, insert_records, check_duplicate, delete_by_source, delete_by_ids, get_recent_history, compute_hash, export_csv)
├── Parser (clean_cell_text, split_items, extract_date_from_text, extract_date_from_filename, find_main_table, parse_table)
├── WordWatcher (폴링 스레드, 문서 파싱 → 콜백)
├── UI - ParseResultPopup (메인 뷰어, 날짜 패널, 테이블, 행 삭제, 저장/스킵)
├── UI - CsvExportDialog (날짜 범위 선택, 파일명 미리보기, 내보내기)
├── UI - ask_date_input (날짜 입력 팝업)
├── Tray (pystray 메뉴, 저장 모드 전환, 토스트)
└── Main (엔트리포인트, 단일 인스턴스, 콜백 연결)
```

테스트 파일:

```
tests/
├── __init__.py
├── test_db.py        (DB CRUD, 해시, CSV 내보내기)
└── test_parser.py    (표 식별, 날짜 추출, 셀 분리, 항목 파싱)
```

---

## Task 1: DB 모듈 — 테스트 + 구현

**Files:**
- Create: `tests/__init__.py`
- Create: `tests/test_db.py`
- Create: `word_crawler.py` (DB 섹션만 먼저)

- [ ] **Step 1: 테스트 파일 생성**

```python
# tests/__init__.py
```

```python
# tests/test_db.py
import os
import csv
import sqlite3
import pytest

# word_crawler.py에서 DB 함수만 import
from word_crawler import (
    init_db, insert_records, check_duplicate, delete_by_source,
    delete_by_ids, get_recent_history, compute_hash, export_csv,
)


@pytest.fixture
def db_path(tmp_path):
    path = str(tmp_path / "test.db")
    init_db(path)
    return path


def _make_record(**overrides):
    base = {
        'date': '2024-05-03', 'source_file': 'test.docx', 'row_num': 2,
        'header1': '구분', 'val1': 'Day',
        'content_col_name': 'A동', 'item_text': 'AHU-3 이상진동',
        'raw_cell': 'AHU-3 이상진동\n\n보일러 점검',
        'header4': '비고', 'val4': '',
    }
    base.update(overrides)
    return base


class TestInitDb:
    def test_creates_table(self, db_path):
        conn = sqlite3.connect(db_path)
        row = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='facility_daily'"
        ).fetchone()
        conn.close()
        assert row is not None

    def test_creates_indexes(self, db_path):
        conn = sqlite3.connect(db_path)
        indexes = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='index' AND name LIKE 'idx_fd_%'"
        ).fetchall()
        conn.close()
        names = [r[0] for r in indexes]
        assert 'idx_fd_date' in names
        assert 'idx_fd_source' in names
        assert 'idx_fd_hash' in names


class TestInsertAndQuery:
    def test_insert_single(self, db_path):
        records = [_make_record()]
        insert_records(db_path, records, content_hash='abc123')
        conn = sqlite3.connect(db_path)
        rows = conn.execute("SELECT * FROM facility_daily").fetchall()
        conn.close()
        assert len(rows) == 1

    def test_insert_multiple(self, db_path):
        records = [
            _make_record(item_text='항목1'),
            _make_record(item_text='항목2'),
            _make_record(item_text='항목3'),
        ]
        insert_records(db_path, records, content_hash='abc123')
        conn = sqlite3.connect(db_path)
        rows = conn.execute("SELECT * FROM facility_daily").fetchall()
        conn.close()
        assert len(rows) == 3


class TestDuplicate:
    def test_new(self, db_path):
        assert check_duplicate(db_path, 'test.docx', '2024-05-03', 'abc') == 'new'

    def test_same(self, db_path):
        insert_records(db_path, [_make_record()], content_hash='abc')
        assert check_duplicate(db_path, 'test.docx', '2024-05-03', 'abc') == 'same'

    def test_changed(self, db_path):
        insert_records(db_path, [_make_record()], content_hash='abc')
        assert check_duplicate(db_path, 'test.docx', '2024-05-03', 'xyz') == 'changed'


class TestDelete:
    def test_delete_by_source(self, db_path):
        insert_records(db_path, [_make_record()], content_hash='abc')
        delete_by_source(db_path, 'test.docx', '2024-05-03')
        conn = sqlite3.connect(db_path)
        rows = conn.execute("SELECT * FROM facility_daily").fetchall()
        conn.close()
        assert len(rows) == 0

    def test_delete_by_ids(self, db_path):
        insert_records(db_path, [_make_record(item_text='a'), _make_record(item_text='b')], content_hash='abc')
        conn = sqlite3.connect(db_path)
        ids = [r[0] for r in conn.execute("SELECT id FROM facility_daily").fetchall()]
        conn.close()
        delete_by_ids(db_path, [ids[0]])
        conn = sqlite3.connect(db_path)
        rows = conn.execute("SELECT * FROM facility_daily").fetchall()
        conn.close()
        assert len(rows) == 1


class TestHistory:
    def test_returns_dicts(self, db_path):
        insert_records(db_path, [_make_record()], content_hash='abc')
        history = get_recent_history(db_path, limit=10)
        assert len(history) == 1
        assert isinstance(history[0], dict)
        assert history[0]['item_text'] == 'AHU-3 이상진동'


class TestHash:
    def test_deterministic(self):
        records = [_make_record()]
        h1 = compute_hash(records)
        h2 = compute_hash(records)
        assert h1 == h2
        assert len(h1) == 16

    def test_different_content_different_hash(self):
        r1 = [_make_record(item_text='a')]
        r2 = [_make_record(item_text='b')]
        assert compute_hash(r1) != compute_hash(r2)


class TestExportCsv:
    def test_export_all(self, db_path, tmp_path):
        insert_records(db_path, [
            _make_record(date='2024-04-01', item_text='항목1'),
            _make_record(date='2024-05-03', item_text='항목2'),
        ], content_hash='abc')
        out = str(tmp_path / "out.csv")
        count = export_csv(db_path, out)
        assert count == 2
        with open(out, encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            header = next(reader)
            assert 'item_text' in header
            rows = list(reader)
            assert len(rows) == 2

    def test_export_date_range(self, db_path, tmp_path):
        insert_records(db_path, [
            _make_record(date='2024-04-01', item_text='항목1'),
            _make_record(date='2024-05-03', item_text='항목2'),
        ], content_hash='abc')
        out = str(tmp_path / "out.csv")
        count = export_csv(db_path, out, start_date='2024-05-01', end_date='2024-05-31')
        assert count == 1
```

- [ ] **Step 2: 테스트 실행 — 실패 확인**

```bash
.venv/Scripts/python -m pytest tests/test_db.py -v
```

Expected: FAIL (import 실패)

- [ ] **Step 3: word_crawler.py DB 섹션 구현**

기존 `word_crawler.py`를 완전히 새로 작성. 이 단계에서는 DB 섹션 + 상수만 포함.

```python
"""
설비일보 Word 크롤러 (단일 파일)
- 열린 Word 문서에서 설비일보 표를 자동 감지/파싱 (동적 헤더)
- SQLite에 적재 (항목 단위 분리 + raw_cell 보존)
- 시스템 트레이 + tkinter 팝업 UI
- CSV 내보내기 → SQream 이관
"""
import os
import sys
import re
import csv
import sqlite3
import hashlib
import threading
import msvcrt

import pythoncom
import win32com.client
import pystray
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import ttk, simpledialog, filedialog

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

DATE_PATTERNS_TEXT = [
    re.compile(r'(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})'),
    re.compile(r'(\d{2})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})'),
]

DATE_PATTERNS_FILENAME = [
    re.compile(r'(\d{4})(\d{2})(\d{2})'),
    re.compile(r'(\d{2})(\d{2})(\d{2})'),
]

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'facility_daily.db')
LOCK_PATH = os.path.join(os.path.dirname(DB_PATH), 'word_crawler.lock')

# ─────────────────────────────────────────────
# DB
# ─────────────────────────────────────────────

def init_db(db_path=None):
    path = db_path or DB_PATH
    os.makedirs(os.path.dirname(path), exist_ok=True)
    conn = sqlite3.connect(path)
    conn.execute('''
        CREATE TABLE IF NOT EXISTS facility_daily (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            date             TEXT NOT NULL,
            source_file      TEXT NOT NULL,
            row_num          INTEGER NOT NULL,
            header1          TEXT,
            val1             TEXT,
            content_col_name TEXT,
            item_text        TEXT,
            raw_cell         TEXT,
            header4          TEXT,
            val4             TEXT,
            content_hash     TEXT,
            created_at       TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_fd_date ON facility_daily(date)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_fd_source ON facility_daily(source_file)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_fd_hash ON facility_daily(content_hash)')
    conn.commit()
    conn.close()


def insert_records(db_path, records, content_hash=None):
    conn = sqlite3.connect(db_path)
    for rec in records:
        conn.execute('''
            INSERT INTO facility_daily
            (date, source_file, row_num, header1, val1, content_col_name,
             item_text, raw_cell, header4, val4, content_hash)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            rec['date'], rec['source_file'], rec['row_num'],
            rec['header1'], rec['val1'], rec['content_col_name'],
            rec['item_text'], rec['raw_cell'],
            rec['header4'], rec['val4'], content_hash,
        ))
    conn.commit()
    conn.close()


def check_duplicate(db_path, source_file, date, new_hash):
    conn = sqlite3.connect(db_path)
    row = conn.execute(
        'SELECT content_hash FROM facility_daily WHERE source_file = ? AND date = ? LIMIT 1',
        (source_file, date)
    ).fetchone()
    conn.close()
    if row is None:
        return 'new'
    return 'same' if row[0] == new_hash else 'changed'


def delete_by_source(db_path, source_file, date):
    conn = sqlite3.connect(db_path)
    conn.execute(
        'DELETE FROM facility_daily WHERE source_file = ? AND date = ?',
        (source_file, date)
    )
    conn.commit()
    conn.close()


def delete_by_ids(db_path, ids):
    if not ids:
        return
    conn = sqlite3.connect(db_path)
    placeholders = ','.join('?' for _ in ids)
    conn.execute(f'DELETE FROM facility_daily WHERE id IN ({placeholders})', ids)
    conn.commit()
    conn.close()


def get_recent_history(db_path, limit=500):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        'SELECT * FROM facility_daily ORDER BY date DESC, id DESC LIMIT ?',
        (limit,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def compute_hash(records):
    raw = '|'.join(
        f"{r.get('val1','')}:{r.get('content_col_name','')}:{r.get('item_text','')}"
        for r in records
    )
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()[:16]


def export_csv(db_path, output_path, start_date=None, end_date=None):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    query = 'SELECT * FROM facility_daily'
    params = []
    conditions = []
    if start_date:
        conditions.append('date >= ?')
        params.append(start_date)
    if end_date:
        conditions.append('date <= ?')
        params.append(end_date)
    if conditions:
        query += ' WHERE ' + ' AND '.join(conditions)
    query += ' ORDER BY date, id'
    rows = conn.execute(query, params).fetchall()
    conn.close()
    if not rows:
        return 0
    keys = rows[0].keys()
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        for row in rows:
            writer.writerow(dict(row))
    return len(rows)
```

나머지 섹션(Parser, WordWatcher, UI, Tray, Main)은 후속 Task에서 이어붙임.

- [ ] **Step 4: 테스트 실행 — 통과 확인**

```bash
.venv/Scripts/python -m pytest tests/test_db.py -v
```

Expected: 모든 테스트 PASS

- [ ] **Step 5: 커밋**

```bash
git add tests/ word_crawler.py
git commit -m "feat: DB 모듈 구현 (새 스키마, CRUD, CSV 내보내기, 행 삭제)"
```

---

## Task 2: Parser 모듈 — 테스트 + 구현

**Files:**
- Create: `tests/test_parser.py`
- Modify: `word_crawler.py` (Parser 섹션 추가)

- [ ] **Step 1: 테스트 파일 생성**

```python
# tests/test_parser.py
import pytest
from word_crawler import (
    clean_cell_text, split_items, extract_date_from_text,
    extract_date_from_filename, find_main_table_index, parse_table_data,
)


class TestCleanCellText:
    def test_removes_control_chars(self):
        raw = 'hello\x07world\x0b\r'
        result = clean_cell_text(raw)
        assert '\x07' not in result
        assert '\r' not in result

    def test_preserves_newlines(self):
        raw = 'line1\nline2'
        result = clean_cell_text(raw)
        assert '\n' in result

    def test_strips(self):
        assert clean_cell_text('  hello  ') == 'hello'


class TestSplitItems:
    def test_single_item(self):
        items = split_items('AHU-3 이상진동 발생')
        assert len(items) == 1
        assert items[0] == 'AHU-3 이상진동 발생'

    def test_double_newline_split(self):
        text = 'AHU-3 이상진동\n\n보일러 점검'
        items = split_items(text)
        assert len(items) == 2
        assert items[0] == 'AHU-3 이상진동'
        assert items[1] == '보일러 점검'

    def test_triple_newline_split(self):
        text = '항목1\n\n\n항목2'
        items = split_items(text)
        assert len(items) == 2

    def test_single_newline_no_split(self):
        text = '한 줄\n이어지는 내용'
        items = split_items(text)
        assert len(items) == 1

    def test_empty(self):
        assert split_items('') == []
        assert split_items('   ') == []


class TestExtractDate:
    def test_from_text_korean(self):
        assert extract_date_from_text('2024년 5월 3일 설비일보') == '2024-05-03'

    def test_from_text_dotted(self):
        assert extract_date_from_text('2024.05.03') == '2024-05-03'

    def test_from_text_dashed(self):
        assert extract_date_from_text('2024-5-3') == '2024-05-03'

    def test_from_text_short_year(self):
        assert extract_date_from_text('24/05/03') == '2024-05-03'

    def test_from_text_none(self):
        assert extract_date_from_text('설비일보') is None

    def test_from_filename_8digit(self):
        assert extract_date_from_filename('설비일보_20240503.docx') == '2024-05-03'

    def test_from_filename_6digit(self):
        assert extract_date_from_filename('설비일보_240503.docx') == '2024-05-03'

    def test_from_filename_none(self):
        assert extract_date_from_filename('설비일보.docx') is None


class TestFindMainTable:
    def test_returns_largest(self):
        # row_counts: [2, 10, 5] → index 1
        assert find_main_table_index([2, 10, 5]) == 1

    def test_single_table(self):
        assert find_main_table_index([5]) == 0

    def test_empty(self):
        assert find_main_table_index([]) is None


class TestParseTableData:
    def test_basic_4col(self):
        headers = ['구분', 'A동', 'B동', '비고']
        rows_data = [
            ['Day', 'AHU-3 이상진동', 'FCU 드레인', '베어링 발주'],
        ]
        records = parse_table_data(headers, rows_data, '2024-05-03', 'test.docx')
        # col2 (A동) 1건 + col3 (B동) 1건 = 2건
        assert len(records) == 2
        a_rec = [r for r in records if r['content_col_name'] == 'A동']
        assert len(a_rec) == 1
        assert a_rec[0]['val1'] == 'Day'
        assert a_rec[0]['item_text'] == 'AHU-3 이상진동'
        assert a_rec[0]['header4'] == '비고'
        assert a_rec[0]['val4'] == '베어링 발주'

    def test_item_split(self):
        headers = ['구분', 'A동', 'B동', '비고']
        rows_data = [
            ['Day', 'AHU 이상\n\n보일러 점검', '특이 없음', ''],
        ]
        records = parse_table_data(headers, rows_data, '2024-05-03', 'test.docx')
        a_recs = [r for r in records if r['content_col_name'] == 'A동']
        assert len(a_recs) == 2
        assert a_recs[0]['item_text'] == 'AHU 이상'
        assert a_recs[1]['item_text'] == '보일러 점검'
        # raw_cell은 동일
        assert a_recs[0]['raw_cell'] == a_recs[1]['raw_cell']

    def test_row_num(self):
        headers = ['구분', 'A동', 'B동', '비고']
        rows_data = [
            ['Day', '항목1', '항목2', ''],
            ['Night', '항목3', '항목4', ''],
        ]
        records = parse_table_data(headers, rows_data, '2024-05-03', 'test.docx')
        day_recs = [r for r in records if r['val1'] == 'Day']
        night_recs = [r for r in records if r['val1'] == 'Night']
        assert all(r['row_num'] == 2 for r in day_recs)
        assert all(r['row_num'] == 3 for r in night_recs)
```

- [ ] **Step 2: 테스트 실행 — 실패 확인**

```bash
.venv/Scripts/python -m pytest tests/test_parser.py -v
```

Expected: FAIL (import 실패)

- [ ] **Step 3: word_crawler.py Parser 섹션 구현**

`word_crawler.py`의 DB 섹션 아래에 추가:

```python
# ─────────────────────────────────────────────
# Parser
# ─────────────────────────────────────────────

def clean_cell_text(raw):
    text = raw.replace('\r\x07', '').replace('\x07', '').replace('\x0b', '\n').replace('\r', '\n')
    text = ''.join(ch if ch == '\n' or (ord(ch) >= 32) else '' for ch in text)
    return text.strip()


def split_items(cell_text):
    text = cell_text.strip()
    if not text:
        return []
    items = re.split(r'\n{2,}', text)
    return [item.strip() for item in items if item.strip()]


def extract_date_from_text(text):
    for pattern in DATE_PATTERNS_TEXT:
        match = pattern.search(text)
        if match:
            groups = match.groups()
            year, month, day = int(groups[0]), int(groups[1]), int(groups[2])
            if year < 100:
                year += 2000
            return f'{year:04d}-{month:02d}-{day:02d}'
    return None


def extract_date_from_filename(filename):
    for pattern in DATE_PATTERNS_FILENAME:
        match = pattern.search(filename)
        if match:
            groups = match.groups()
            if len(groups) == 3:
                year, month, day = int(groups[0]), int(groups[1]), int(groups[2])
                if year < 100:
                    year += 2000
                return f'{year:04d}-{month:02d}-{day:02d}'
    return None


def find_main_table_index(row_counts):
    if not row_counts:
        return None
    return row_counts.index(max(row_counts))


def parse_table_data(headers, rows_data, date_str, source_file):
    records = []
    h1_name = headers[0] if len(headers) > 0 else ''
    h4_name = headers[3] if len(headers) > 3 else ''
    content_cols = []
    if len(headers) > 1:
        content_cols.append((1, headers[1]))
    if len(headers) > 2:
        content_cols.append((2, headers[2]))

    for row_idx, row in enumerate(rows_data):
        row_num = row_idx + 2  # 1행은 헤더, 데이터는 2행부터
        val1 = row[0] if len(row) > 0 else ''
        val4 = row[3] if len(row) > 3 else ''

        for col_idx, col_name in content_cols:
            if col_idx >= len(row):
                continue
            raw_cell = row[col_idx]
            items = split_items(raw_cell)
            if not items:
                items = [raw_cell] if raw_cell.strip() else []
            for item_text in items:
                records.append({
                    'date': date_str,
                    'source_file': source_file,
                    'row_num': row_num,
                    'header1': h1_name,
                    'val1': val1,
                    'content_col_name': col_name,
                    'item_text': item_text,
                    'raw_cell': raw_cell,
                    'header4': h4_name,
                    'val4': val4,
                })

    return records
```

- [ ] **Step 4: 테스트 실행 — 통과 확인**

```bash
.venv/Scripts/python -m pytest tests/test_parser.py -v
```

Expected: 모든 테스트 PASS

- [ ] **Step 5: 커밋**

```bash
git add tests/test_parser.py word_crawler.py
git commit -m "feat: Parser 모듈 구현 (동적 헤더, 항목 분리, 날짜 추출)"
```

---

## Task 3: WordWatcher 모듈 구현

**Files:**
- Modify: `word_crawler.py` (WordWatcher 섹션 추가)

- [ ] **Step 1: word_crawler.py에 WordWatcher 클래스 추가**

Parser 섹션 아래에 추가:

```python
# ─────────────────────────────────────────────
# Word Watcher
# ─────────────────────────────────────────────

class WordWatcher:
    def __init__(self, db_path, on_new_parse, on_duplicate_same,
                 on_duplicate_changed, on_date_missing, on_no_table):
        self.db_path = db_path
        self.on_new_parse = on_new_parse
        self.on_duplicate_same = on_duplicate_same
        self.on_duplicate_changed = on_duplicate_changed
        self.on_date_missing = on_date_missing
        self.on_no_table = on_no_table
        self._stop_event = threading.Event()
        self._seen_docs = set()
        self._thread = None

    def start(self):
        self._thread = threading.Thread(target=self._watch_loop, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop_event.set()

    def _watch_loop(self):
        pythoncom.CoInitialize()
        try:
            while not self._stop_event.is_set():
                try:
                    self._check_word()
                except Exception:
                    pass
                self._stop_event.wait(4)
        finally:
            pythoncom.CoUninitialize()

    def _check_word(self):
        try:
            word = win32com.client.GetActiveObject('Word.Application')
        except Exception:
            return

        for i in range(1, word.Documents.Count + 1):
            doc = word.Documents(i)
            doc_full = doc.FullName
            if doc_full in self._seen_docs:
                continue

            result = self._parse_document(doc)
            self._seen_docs.add(doc_full)

            if result is None:
                self.on_no_table(doc.Name)
                continue

            records, date_str = result
            content_hash = compute_hash(records)
            dup_status = check_duplicate(self.db_path, doc.Name, date_str, content_hash)

            if dup_status == 'new':
                self.on_new_parse(doc.Name, date_str, records, content_hash)
            elif dup_status == 'same':
                self.on_duplicate_same(doc.Name, date_str)
            else:
                self.on_duplicate_changed(doc.Name, date_str, records, content_hash)

    def _parse_document(self, doc):
        table_count = doc.Tables.Count
        if table_count == 0:
            return None

        # 각 표의 행 수 수집 → 가장 큰 표 선택
        row_counts = []
        for t in range(1, table_count + 1):
            try:
                row_counts.append(doc.Tables(t).Rows.Count)
            except Exception:
                row_counts.append(0)

        table_idx = find_main_table_index(row_counts)
        if table_idx is None:
            return None

        table = doc.Tables(table_idx + 1)  # COM은 1-indexed

        # 헤더 동적 읽기 (1행)
        headers = []
        row1 = table.Rows(1)
        for c in range(1, row1.Cells.Count + 1):
            h_raw = row1.Cells(c).Range.Text
            headers.append(clean_cell_text(h_raw))

        # 날짜 추출
        date_str = None
        table_range = table.Range
        doc_text_before = doc.Range(0, table_range.Start).Text
        date_str = extract_date_from_text(doc_text_before)
        if date_str is None:
            date_str = extract_date_from_filename(doc.Name)
        if date_str is None:
            date_str = self.on_date_missing(doc.Name)
            if date_str is None:
                return None

        # 데이터 행 읽기
        rows_data = []
        for r in range(2, table.Rows.Count + 1):
            row = table.Rows(r)
            cells = []
            for c in range(1, row.Cells.Count + 1):
                raw = row.Cells(c).Range.Text
                cells.append(clean_cell_text(raw))
            rows_data.append(cells)

        records = parse_table_data(headers, rows_data, date_str, doc.Name)
        return (records, date_str) if records else None

    def reset_seen(self):
        self._seen_docs.clear()
```

- [ ] **Step 2: 커밋**

```bash
git add word_crawler.py
git commit -m "feat: WordWatcher 구현 (동적 표 식별, COM 폴링)"
```

---

## Task 4: UI — ParseResultPopup + CsvExportDialog + ask_date_input

**Files:**
- Modify: `word_crawler.py` (UI 섹션 추가)

- [ ] **Step 1: UI 전체 코드 추가**

WordWatcher 섹션 아래에 추가. 목업(`mockup_word_crawler_ui.html`) 구조를 따르는 tkinter 구현:

- `ParseResultPopup`: 날짜 패널 + 테이블 + 저장/스킵/행 삭제
- `CsvExportDialog`: 날짜 범위 선택 + 파일명 미리보기 + 내보내기
- `ask_date_input`: 날짜 입력 simpledialog

핵심 변경점 (기존 대비):
- 테이블 컬럼: 날짜 / 구분(val1) / 영역(content_col_name) / 항목(item_text) / 원문(raw_cell) / 비고(val4)
- **행 삭제**: 선택된 행의 DB id로 `delete_by_ids` 호출 (저장 후) 또는 pending에서 제거 (저장 전)
- **CSV 내보내기 버튼**: 팝업 하단 또는 트레이 메뉴에서 `CsvExportDialog` 호출

전체 코드는 기존 UI 섹션(약 200줄)을 새 스키마 + 행 삭제 + CSV 다이얼로그로 확장.

- [ ] **Step 2: 수동 UI 테스트**

```bash
.venv/Scripts/python -c "
import tkinter as tk
from word_crawler import ParseResultPopup, init_db, DB_PATH
init_db()
root = tk.Tk()
root.withdraw()
popup = ParseResultPopup.get_or_create(on_save_all=lambda p: print('저장'), db_path=DB_PATH)
popup.add_records('test.docx', '2024-05-03', [
    {'date':'2024-05-03','source_file':'test.docx','row_num':2,'header1':'구분','val1':'Day',
     'content_col_name':'A동','item_text':'AHU-3 테스트','raw_cell':'AHU-3 테스트','header4':'비고','val4':''},
], 'hash123')
"
```

Expected: 팝업 열림, 테이블에 1건 표시

- [ ] **Step 3: 커밋**

```bash
git add word_crawler.py
git commit -m "feat: UI 구현 (파싱 뷰어, 행 삭제, CSV 내보내기 다이얼로그)"
```

---

## Task 5: Tray + Main (저장 모드 전환 포함)

**Files:**
- Modify: `word_crawler.py` (Tray + Main 섹션 추가)

- [ ] **Step 1: Tray 클래스 구현**

트레이 메뉴:
- 파싱 뷰어 열기
- (구분선)
- ✓ 확인 후 저장 / 즉시 저장 (라디오 선택)
- (구분선)
- 종료

```python
class TrayApp:
    def __init__(self, db_path, save_all_callback=None):
        self.db_path = db_path
        self.save_all_callback = save_all_callback or (lambda p: None)
        self.icon = None
        self.auto_save = False  # False = 확인 후 저장 (기본)

    def start(self):
        menu = pystray.Menu(
            pystray.MenuItem('파싱 뷰어 열기', self._show_viewer),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                '확인 후 저장',
                self._set_confirm_mode,
                checked=lambda item: not self.auto_save
            ),
            pystray.MenuItem(
                '즉시 저장',
                self._set_auto_mode,
                checked=lambda item: self.auto_save
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('종료', self._quit),
        )
        self.icon = pystray.Icon(
            '설비일보 Word 파서',
            create_icon_image(),
            '설비일보 Word 파서',
            menu
        )
        self.icon.run()

    # ... (notify, _show_viewer, _set_confirm_mode, _set_auto_mode, _quit)
```

- [ ] **Step 2: Main 함수 구현**

콜백 연결:
- `on_new_parse`: auto_save 모드면 즉시 저장 + 토스트, 아니면 팝업
- `on_duplicate_same`: 토스트만
- `on_duplicate_changed`: 팝업
- `on_no_table`: "파싱 대상 아님" 토스트
- `on_date_missing`: `ask_date_input`

- [ ] **Step 3: 통합 테스트**

```bash
.venv/Scripts/python word_crawler.py
```

1. 트레이 아이콘 표시 확인
2. Word에서 더미 docx 열기
3. 파싱 결과 팝업 확인
4. 저장 → SQLite 확인
5. 트레이 우클릭 → 모드 전환 확인
6. CSV 내보내기 확인

- [ ] **Step 4: 커밋**

```bash
git add word_crawler.py
git commit -m "feat: Tray + Main 구현 (저장 모드 전환, 콜백 연결)"
```

---

## Task 6: 전체 테스트 + 정리

**Files:**
- Modify: `word_crawler.py` (최종 정리)

- [ ] **Step 1: 전체 테스트 실행**

```bash
.venv/Scripts/python -m pytest tests/ -v
```

Expected: 모든 테스트 PASS

- [ ] **Step 2: 기존 파일 정리**

- 기존 `excel_crawler.py`는 그대로 유지 (별도 프로젝트)
- `mockup_word_crawler_ui.html`, 스크린샷 파일은 참조용으로 유지

- [ ] **Step 3: 최종 커밋**

```bash
git add -A
git commit -m "feat: 설비일보 Word 크롤러 v2.0 완성 (동적 파싱, 항목 분리, CSV 내보내기)"
```
