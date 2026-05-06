"""
설비일보 Word 크롤러 (단일 파일)
- 열린 Word 문서에서 설비일보 표를 자동 감지/파싱 (동적 헤더)
- *제목 + 본문 형식 항목 추출 (2+ 개행으로 항목 분리)
- SQLite 적재 + CSV 내보내기 → SQream 이관
- 시스템 트레이 + tkinter 다크 테마 팝업 UI

필요 라이브러리:
    pip install pystray pywin32 Pillow

빌드 (단일 exe):
    pip install pyinstaller
    pyinstaller --onefile --noconsole --name word_crawler word_crawler.py

실행:
    python word_crawler.py
    (빌드 후) word_crawler.exe

파일 구조 (실행 시 자동 생성):
    word_crawler.exe (또는 .py)
    autosave/
    ├── facility_daily.db    ← SQLite DB
    ├── word_crawler.lock    ← 단일 인스턴스 락
    └── *.csv                ← 내보낸 CSV 파일
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
from tkinter import ttk, simpledialog, filedialog, messagebox

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

DATE_PATTERNS_TEXT = [
    # "일자 : 2024년 5월 3일" 형식 우선
    re.compile(r'일\s*자[:：\s]+(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일'),
    re.compile(r'(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})'),
    re.compile(r'(\d{2})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})'),
]

DATE_PATTERNS_FILENAME = [
    re.compile(r'(\d{4})(\d{2})(\d{2})'),
    re.compile(r'(\d{2})(\d{2})(\d{2})'),
]

# PyInstaller 빌드 시 exe 경로, 일반 실행 시 py 경로
if getattr(sys, 'frozen', False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))

AUTOSAVE_DIR = os.path.join(_BASE_DIR, 'autosave')
DB_PATH = os.path.join(AUTOSAVE_DIR, 'facility_daily.db')
LOCK_PATH = os.path.join(AUTOSAVE_DIR, 'word_crawler.lock')

# 다크 테마 색상
C = {
    'bg_deep': '#0c111b',
    'bg_panel': '#131a2a',
    'bg_surface': '#1a2236',
    'bg_elevated': '#212d45',
    'border': '#2a3650',
    'accent': '#3b82f6',
    'green': '#22c55e',
    'green_row': '#0f1f15',
    'amber': '#f59e0b',
    'red': '#ef4444',
    'text1': '#e2e8f0',
    'text2': '#94a3b8',
    'text3': '#64748b',
    'text_accent': '#93c5fd',
}

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
            title            TEXT,
            raw_text         TEXT,
            raw_cell         TEXT,
            header4          TEXT,
            val4             TEXT,
            content_hash     TEXT,
            created_at       TEXT DEFAULT (datetime('now', 'localtime')),
            exported_at      TEXT
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_fd_date ON facility_daily(date)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_fd_source ON facility_daily(source_file)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_fd_hash ON facility_daily(content_hash)')
    # 기존 DB 마이그레이션: exported_at 컬럼 없으면 추가
    try:
        conn.execute('ALTER TABLE facility_daily ADD COLUMN exported_at TEXT')
    except Exception:
        pass
    # 기존 DB 마이그레이션: phenomenon/cause/action 컬럼 제거 (SQLite 3.35+)
    for col in ('phenomenon', 'cause', 'action'):
        try:
            conn.execute(f'ALTER TABLE facility_daily DROP COLUMN {col}')
        except Exception:
            pass
    conn.commit()
    conn.close()


def insert_records(db_path, records, content_hash=None):
    conn = sqlite3.connect(db_path)
    for rec in records:
        conn.execute('''
            INSERT INTO facility_daily
            (date, source_file, row_num, header1, val1, content_col_name,
             title, raw_text, raw_cell, header4, val4, content_hash)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            rec['date'], rec['source_file'], rec['row_num'],
            rec['header1'], rec['val1'], rec['content_col_name'],
            rec.get('title', ''), rec.get('raw_text', ''), rec['raw_cell'],
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
    conn.execute('DELETE FROM facility_daily WHERE source_file = ? AND date = ?',
                 (source_file, date))
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
        'SELECT * FROM facility_daily ORDER BY date DESC, id DESC LIMIT ?', (limit,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def compute_hash(records):
    raw = '|'.join(
        f"{r.get('val1','')}:{r.get('content_col_name','')}:{r.get('raw_text','')}"
        for r in records
    )
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()[:16]


def export_csv(db_path, output_path, start_date=None, end_date=None):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    query = 'SELECT * FROM facility_daily'
    params, conds = [], []
    if start_date:
        conds.append('date >= ?'); params.append(start_date)
    if end_date:
        conds.append('date <= ?'); params.append(end_date)
    if conds:
        query += ' WHERE ' + ' AND '.join(conds)
    query += ' ORDER BY date, id'
    rows = conn.execute(query, params).fetchall()
    if not rows:
        conn.close()
        return 0
    keys = rows[0].keys()
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        for row in rows:
            writer.writerow(dict(row))
    # 내보낸 레코드에 exported_at 기록
    ids = [row['id'] for row in rows]
    placeholders = ','.join('?' for _ in ids)
    conn.execute(
        f"UPDATE facility_daily SET exported_at = datetime('now','localtime') WHERE id IN ({placeholders})",
        ids)
    conn.commit()
    conn.close()
    return len(rows)


def get_export_status_by_date(db_path):
    """날짜별 내보내기 상태 반환. {date: 'all'|'partial'|'none'}"""
    conn = sqlite3.connect(db_path)
    rows = conn.execute('''
        SELECT date,
               COUNT(*) as total,
               COUNT(exported_at) as exported
        FROM facility_daily GROUP BY date
    ''').fetchall()
    conn.close()
    result = {}
    for date, total, exported in rows:
        if exported == 0:
            result[date] = 'none'
        elif exported >= total:
            result[date] = 'all'
        else:
            result[date] = 'partial'
    return result

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


def _strip_numbering(text):
    """각 줄 앞의 번호 패턴(1), 2), 3) 등) 제거."""
    return re.sub(r'^\d+\)\s*', '', text.strip(), flags=re.MULTILINE)


_TITLE_LINE_RE = re.compile(r'^\*\s*(.+)$')


def parse_item_block(text):
    """한 블록(2+ 개행으로 분리된 단위)을 title + raw_text 로 분리.

    포맷:
        *제목
         - 내용 1
         - 내용 2

    `*` 로 시작하는 첫 줄은 title 로 추출하고 나머지 줄은 raw_text 에 이어
    붙인다. `*` 라인이 없으면 title 은 비우고 전체 텍스트를 raw_text 에 둔다.
    1) 2) 형식 번호 접두는 제거한다."""
    text = _strip_numbering(text)
    parsed = {'title': '', 'raw_text': ''}

    body_lines = []
    for line in text.split('\n'):
        stripped = line.strip()
        if not stripped:
            continue
        m = _TITLE_LINE_RE.match(stripped)
        if m and not parsed['title']:
            parsed['title'] = m.group(1).strip()
            continue
        body_lines.append(stripped)

    parsed['raw_text'] = '\n'.join(body_lines)
    return parsed


def extract_date_from_text(text):
    for pattern in DATE_PATTERNS_TEXT:
        match = pattern.search(text)
        if match:
            groups = match.groups()
            y, m, d = int(groups[0]), int(groups[1]), int(groups[2])
            if y < 100:
                y += 2000
            return f'{y:04d}-{m:02d}-{d:02d}'
    return None


def extract_date_from_filename(filename):
    for pattern in DATE_PATTERNS_FILENAME:
        match = pattern.search(filename)
        if match:
            groups = match.groups()
            if len(groups) == 3:
                y, m, d = int(groups[0]), int(groups[1]), int(groups[2])
                if y < 100:
                    y += 2000
                return f'{y:04d}-{m:02d}-{d:02d}'
    return None


REQUIRED_HEADERS = ('구분', 'UT동', '확산동', '전달사항')


def format_multiline(text, sep=' | '):
    """여러 줄 텍스트를 단일 줄 + 명시 구분자로 변환. 뷰어/CSV 출력에서
    개행 구조가 사라지지 않도록 라인 사이에 ` | ` 같은 마커를 끼워 넣는다.
    빈 줄은 제거한다."""
    if not text:
        return ''
    s = str(text).replace('\r\n', '\n').replace('\r', '\n')
    lines = [line.strip() for line in s.split('\n') if line.strip()]
    return sep.join(lines)


def find_main_table_index(headers_list):
    """각 표의 1행 헤더 리스트(list[list[str]])를 받아 메인 표 인덱스를 반환.

    REQUIRED_HEADERS(`구분`, `UT동`, `확산동`, `전달사항`) 가 1행에 모두
    포함된 첫 번째 표를 메인 표로 선택. 매칭되는 표가 없으면 None."""
    if not headers_list:
        return None
    for idx, headers in enumerate(headers_list):
        normalized = [re.sub(r'\s+', '', h or '') for h in headers]
        if all(req in normalized for req in REQUIRED_HEADERS):
            return idx
    return None


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
        row_num = row_idx + 2
        val1 = row[0] if len(row) > 0 else ''
        val4 = row[3] if len(row) > 3 else ''

        for col_idx, col_name in content_cols:
            if col_idx >= len(row):
                continue
            raw_cell = row[col_idx]
            items = split_items(raw_cell)
            if not items:
                items = [raw_cell] if raw_cell.strip() else []
            for block_text in items:
                parsed = parse_item_block(block_text)
                records.append({
                    'date': date_str,
                    'source_file': source_file,
                    'row_num': row_num,
                    'header1': h1_name,
                    'val1': val1,
                    'content_col_name': col_name,
                    'title': parsed['title'],
                    'raw_text': parsed['raw_text'],
                    'raw_cell': raw_cell,
                    'header4': h4_name,
                    'val4': val4,
                })

    return records

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

        # 현재 열려있는 문서 목록 수집
        current_docs = set()
        for i in range(1, word.Documents.Count + 1):
            try:
                current_docs.add(word.Documents(i).FullName)
            except Exception:
                pass

        # 닫힌 문서를 _seen_docs에서 제거 → 다시 열면 재파싱 가능
        self._seen_docs -= (self._seen_docs - current_docs)

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
            dup = check_duplicate(self.db_path, doc.Name, date_str, content_hash)
            if dup == 'new':
                self.on_new_parse(doc.Name, date_str, records, content_hash)
            elif dup == 'same':
                self.on_duplicate_same(doc.Name, date_str)
            else:
                self.on_duplicate_changed(doc.Name, date_str, records, content_hash)

    def _parse_document(self, doc):
        if doc.Tables.Count == 0:
            return None
        # 각 표의 1행 헤더를 수집해 REQUIRED_HEADERS 매칭으로 메인 표 선정
        headers_list = []
        for t in range(1, doc.Tables.Count + 1):
            try:
                row1 = doc.Tables(t).Rows(1)
                cells = [clean_cell_text(row1.Cells(c).Range.Text)
                         for c in range(1, row1.Cells.Count + 1)]
                headers_list.append(cells)
            except Exception:
                headers_list.append([])
        table_idx = find_main_table_index(headers_list)
        if table_idx is None:
            return None
        table = doc.Tables(table_idx + 1)
        headers = headers_list[table_idx]
        # 본문 + 텍스트 박스(Shape) + 헤더/푸터 텍스트를 모아서 날짜 검색
        date_str = extract_date_from_text(self._collect_text_for_date(doc, table))
        if date_str is None:
            date_str = extract_date_from_filename(doc.Name)
        if date_str is None:
            date_str = self.on_date_missing(doc.Name)
            if date_str is None:
                return None
        rows_data = []
        for r in range(2, table.Rows.Count + 1):
            row = table.Rows(r)
            cells = []
            for c in range(1, row.Cells.Count + 1):
                cells.append(clean_cell_text(row.Cells(c).Range.Text))
            rows_data.append(cells)
        records = parse_table_data(headers, rows_data, date_str, doc.Name)
        return (records, date_str) if records else None

    @staticmethod
    def _collect_text_for_date(doc, table):
        """날짜 추출용 텍스트 묶음. 표 위 본문 + 모든 Shape(텍스트 박스) +
        헤더/푸터 텍스트를 합쳐 반환. 일부 양식은 날짜를 본문 텍스트가 아닌
        텍스트 박스에 그려두기 때문에 Shapes 까지 훑어야 함."""
        parts = []
        # 본문에서 표 위쪽 텍스트
        try:
            parts.append(doc.Range(0, table.Range.Start).Text)
        except Exception:
            pass
        # 떠 있는 Shape (텍스트 박스 포함)
        try:
            for i in range(1, doc.Shapes.Count + 1):
                try:
                    shape = doc.Shapes(i)
                    tf = shape.TextFrame
                    if tf.HasText:
                        parts.append(tf.TextRange.Text)
                except Exception:
                    continue
        except Exception:
            pass
        # 인라인 Shape (드물게 텍스트 박스가 인라인일 수 있음)
        try:
            for i in range(1, doc.InlineShapes.Count + 1):
                try:
                    ish = doc.InlineShapes(i)
                    tr = ish.Range
                    parts.append(tr.Text)
                except Exception:
                    continue
        except Exception:
            pass
        # 섹션별 헤더/푸터
        try:
            for s in range(1, doc.Sections.Count + 1):
                section = doc.Sections(s)
                for hf in (section.Headers, section.Footers):
                    try:
                        for j in range(1, hf.Count + 1):
                            try:
                                parts.append(hf(j).Range.Text)
                            except Exception:
                                continue
                    except Exception:
                        continue
        except Exception:
            pass
        return '\n'.join(p for p in parts if p)

    def reset_seen(self):
        self._seen_docs.clear()

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────

MULTILINE_TREE_COLS = {'title', 'raw_text', 'raw_cell'}


def _flat(text):
    """뷰어용 단일 라인 변환. 멀티라인은 ` | ` 구분자로 합친다."""
    if not text:
        return ''
    return format_multiline(str(text))


def _tree_cell(rec, col_id):
    """Treeview 셀 표시용. 모든 컬럼을 한 줄로 변환. 멀티라인 컬럼은
    ` | ` 구분자로 합쳐 한 줄에 들어가게 한다(원본 \\n 은 record dict 에
    그대로 보존되어 더블클릭 팝업에서 사용)."""
    val = rec.get(col_id, '')
    if not val:
        return ''
    return format_multiline(str(val))


def _apply_dark_theme(root):
    root.configure(bg=C['bg_deep'])
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('Treeview',
                    background=C['bg_panel'], foreground=C['text1'],
                    fieldbackground=C['bg_panel'], borderwidth=0,
                    font=('맑은 고딕', 9), rowheight=24,
                    anchor='center')
    style.configure('Treeview.Heading',
                    background=C['bg_surface'], foreground=C['text_accent'],
                    font=('맑은 고딕', 9, 'bold'), borderwidth=0,
                    anchor='center')
    style.map('Treeview',
              background=[('selected', C['accent'])],
              foreground=[('selected', 'white')])
    style.map('Treeview.Heading',
              background=[('active', C['bg_elevated'])])
    style.configure('Vertical.TScrollbar',
                    background=C['bg_surface'], troughcolor=C['bg_deep'],
                    borderwidth=0, arrowsize=12)
    style.configure('Horizontal.TScrollbar',
                    background=C['bg_surface'], troughcolor=C['bg_deep'],
                    borderwidth=0, arrowsize=12)


def _dark_frame(parent, **kw):
    return tk.Frame(parent, bg=C['bg_deep'], **kw)


def _dark_label(parent, text='', **kw):
    return tk.Label(parent, text=text, bg=C['bg_deep'], fg=C['text1'],
                    font=('맑은 고딕', 10), **kw)


def _dark_button(parent, text='', command=None, style='default', **kw):
    colors = {
        'default': {'bg': C['bg_elevated'], 'fg': C['text2'], 'abg': C['border']},
        'primary': {'bg': C['accent'], 'fg': 'white', 'abg': '#2563eb'},
        'danger': {'bg': '#7f1d1d', 'fg': '#fca5a5', 'abg': C['red']},
        'export': {'bg': '#14532d', 'fg': '#86efac', 'abg': C['green']},
    }
    c = colors.get(style, colors['default'])
    return tk.Button(parent, text=text, command=command,
                     bg=c['bg'], fg=c['fg'], activebackground=c['abg'],
                     activeforeground='white', relief='flat', bd=0,
                     font=('맑은 고딕', 10), padx=16, pady=5, cursor='hand2', **kw)


class ParseResultPopup:
    _instance = None
    _cls_lock = threading.Lock()

    @classmethod
    def get_or_create(cls, on_save_all=None, db_path=None):
        with cls._cls_lock:
            if cls._instance is None or not cls._instance._alive:
                cls._instance = cls(on_save_all=on_save_all, db_path=db_path)
            return cls._instance

    def __init__(self, on_save_all=None, db_path=None):
        self.on_save_all = on_save_all or (lambda p: None)
        self.db_path = db_path
        self._alive = False
        self._pending = []
        self._lock = threading.Lock()
        self._root = None
        self._tree = None
        self._row_records = {}  # Treeview item id → 원본 record dict (개행 포함)
        self._date_listbox = None
        self._info_var = None
        self._db_records = []
        self._current_filter_date = None

    def add_records(self, doc_name, date_str, records, content_hash):
        with self._lock:
            self._pending.append((doc_name, date_str, records, content_hash))
        if self._root and self._alive:
            self._root.after(0, self._refresh_tree)
        else:
            threading.Thread(target=self._show, daemon=True).start()

    def _show(self):
        self._alive = True
        root = tk.Tk()
        self._root = root
        root.title('설비일보 파서 (Word)')
        root.geometry('1300x650')
        root.resizable(True, True)
        root.protocol('WM_DELETE_WINDOW', self._on_close)
        _apply_dark_theme(root)

        # ── 상단 정보 바 ──
        top_bar = tk.Frame(root, bg=C['bg_surface'], height=36)
        top_bar.pack(fill='x')
        top_bar.pack_propagate(False)
        self._info_var = tk.StringVar(value='')
        tk.Label(top_bar, textvariable=self._info_var,
                 bg=C['bg_surface'], fg=C['text2'],
                 font=('맑은 고딕', 10), padx=12).pack(side='left', fill='y')
        self._status_label = tk.Label(top_bar, text='● 감지 중',
                                       bg=C['bg_surface'], fg=C['green'],
                                       font=('맑은 고딕', 9), padx=12)
        self._status_label.pack(side='right', fill='y')

        # ── 메인 영역 ──
        main = _dark_frame(root)
        main.pack(fill='both', expand=True, padx=8, pady=4)

        # 좌측 날짜 패널
        left = tk.Frame(main, bg=C['bg_panel'], width=155,
                        highlightbackground=C['border'], highlightthickness=1)
        left.pack(side='left', fill='y', padx=(0, 6))
        left.pack_propagate(False)

        tk.Label(left, text='날짜 목록', bg=C['bg_panel'], fg=C['text3'],
                 font=('맑은 고딕', 9, 'bold'), anchor='w', padx=8).pack(fill='x', pady=(8, 4))

        self._date_listbox = tk.Listbox(left, selectmode='single',
                                        font=('맑은 고딕', 9),
                                        bg=C['bg_deep'], fg=C['text2'],
                                        selectbackground=C['accent'], selectforeground='white',
                                        highlightthickness=0, bd=0, relief='flat')
        self._date_listbox.pack(fill='both', expand=True, padx=4)
        self._date_listbox.bind('<<ListboxSelect>>', self._on_date_select)

        _dark_button(left, '전체 보기', self._show_all_dates).pack(
            fill='x', padx=4, pady=6)

        # 우측 Treeview
        right = _dark_frame(main)
        right.pack(side='left', fill='both', expand=True)

        # (db_col_id, db_col_name, display_label, width)
        self._col_defs = [
            ('date',             'date',             '날짜',   85),
            ('source_file',      'source_file',      '파일명', 120),
            ('row_num',          'row_num',           '행번호', 50),
            ('header1',          'header1',           '헤더1',  60),
            ('val1',             'val1',              '구분',   60),
            ('content_col_name', 'content_col_name',  '영역',   60),
            ('title',            'title',             '제목',  140),
            ('raw_text',         'raw_text',          '내용',  240),
            ('raw_cell',         'raw_cell',          '원문',  150),
            ('header4',          'header4',           '헤더4',  60),
            ('val4',             'val4',              '비고',  100),
            ('content_hash',     'content_hash',      '해시',   80),
        ]
        cols = tuple(d[0] for d in self._col_defs)

        tree_frame = _dark_frame(right)
        tree_frame.pack(fill='both', expand=True)

        # ── 고정 한글 헤더 Treeview (스크롤 안 됨) ──
        self._header_tree = ttk.Treeview(tree_frame, columns=cols, show='headings',
                                          height=1, selectmode='none')
        for col_id, db_name, label, w in self._col_defs:
            self._header_tree.heading(col_id, text=db_name, anchor='center')
            self._header_tree.column(col_id, width=w, minwidth=40, stretch=False, anchor='center')
        self._header_tree.tag_configure('label_row',
                                         background=C['bg_surface'], foreground=C['text_accent'])
        display_labels = tuple(d[2] for d in self._col_defs)
        self._header_tree.insert('', 'end', values=display_labels, tags=('label_row',))
        self._header_tree.pack(fill='x')
        # 고정 헤더 클릭/선택 방지
        self._header_tree.bind('<Button-1>', lambda e: 'break')
        self._header_tree.bind('<Double-1>', lambda e: 'break')

        # ── 데이터 Treeview (스크롤 가능) ──
        data_frame = _dark_frame(tree_frame)
        data_frame.pack(fill='both', expand=True)

        sy = ttk.Scrollbar(data_frame, orient='vertical')
        sx = ttk.Scrollbar(tree_frame, orient='horizontal')

        # 데이터 Treeview 전용 스타일: 멀티라인 셀에 맞춰 rowheight 동적 조정
        data_style = ttk.Style(root)
        data_style.configure('Data.Treeview',
                              background=C['bg_panel'], foreground=C['text1'],
                              fieldbackground=C['bg_panel'], borderwidth=0,
                              font=('맑은 고딕', 9), rowheight=24)
        data_style.map('Data.Treeview',
                        background=[('selected', C['accent'])],
                        foreground=[('selected', 'white')])
        self._tree = ttk.Treeview(data_frame, columns=cols, show='headings',
                                   yscrollcommand=sy.set, xscrollcommand=sx.set,
                                   selectmode='extended', style='Data.Treeview')
        sy.config(command=self._tree.yview)

        # 가로 스크롤을 두 Treeview 동기화
        def _sync_xscroll(*args):
            self._tree.xview(*args)
            self._header_tree.xview(*args)
        sx.config(command=_sync_xscroll)
        self._tree.configure(xscrollcommand=sx.set)

        for col_id, db_name, label, w in self._col_defs:
            self._tree.heading(col_id, text='')  # 데이터 Treeview 헤더는 숨김 (고정 헤더가 대체)
            self._tree.column(col_id, width=w, minwidth=40, stretch=False)

        # 데이터 Treeview 네이티브 헤더 숨김
        style = ttk.Style(root)
        style.configure('NoHeader.Treeview.Heading', background=C['bg_deep'],
                        foreground=C['bg_deep'], borderwidth=0, relief='flat')
        style.layout('NoHeader.Treeview', style.layout('Treeview'))
        # Treeview 헤더 높이를 0으로 만들 수 없으므로 show='tree headings' 대신 비움
        self._tree.configure(show='')  # 헤더 완전 숨김

        self._tree.tag_configure('new', background=C['green_row'])
        self._tree.bind('<Double-1>', self._on_cell_double_click)

        sy.pack(side='right', fill='y')
        self._tree.pack(fill='both', expand=True)
        sx.pack(fill='x')

        # ── 하단 버튼 바 ──
        bot = tk.Frame(root, bg=C['bg_surface'], height=50)
        bot.pack(fill='x')
        bot.pack_propagate(False)

        _dark_button(bot, '선택 행 삭제', self._delete_selected, 'danger').pack(
            side='left', padx=12, pady=8)

        _dark_button(bot, '저장', self._save_all, 'primary').pack(
            side='right', padx=12, pady=8)
        _dark_button(bot, '스킵', self._on_close).pack(
            side='right', pady=8)

        # ── CSV 내보내기 패널 (하단 인라인) ──
        csv_panel = tk.Frame(root, bg=C['bg_panel'],
                             highlightbackground=C['border'], highlightthickness=1)
        csv_panel.pack(fill='x', padx=8, pady=(0, 8))

        csv_header = tk.Frame(csv_panel, bg=C['bg_surface'])
        csv_header.pack(fill='x')
        tk.Label(csv_header, text='CSV 내보내기', font=('맑은 고딕', 10, 'bold'),
                 bg=C['bg_surface'], fg=C['text_accent'], padx=12, pady=6).pack(side='left')

        csv_body = tk.Frame(csv_panel, bg=C['bg_panel'])
        csv_body.pack(fill='x', padx=12, pady=8)

        # 범위 선택
        tk.Label(csv_body, text='범위', bg=C['bg_panel'], fg=C['text3'],
                 font=('맑은 고딕', 9)).grid(row=0, column=0, sticky='w', pady=2)
        self._csv_range_var = tk.StringVar(value='날짜 범위 지정')
        range_combo = ttk.Combobox(csv_body, textvariable=self._csv_range_var,
                                    values=['날짜 범위 지정', '전체 내보내기'],
                                    state='readonly', width=16)
        range_combo.grid(row=0, column=1, padx=(8, 16), pady=2, sticky='w')
        range_combo.bind('<<ComboboxSelected>>', self._csv_on_range_change)

        # 시작일 (Combobox — DB 날짜 자동 파악)
        tk.Label(csv_body, text='시작일', bg=C['bg_panel'], fg=C['text3'],
                 font=('맑은 고딕', 9)).grid(row=0, column=2, sticky='w', pady=2)
        self._csv_start_var = tk.StringVar()
        self._csv_start_combo = ttk.Combobox(csv_body, textvariable=self._csv_start_var,
                 font=('맑은 고딕', 10), width=12, state='readonly')
        self._csv_start_combo.grid(row=0, column=3, padx=(4, 16), pady=2)

        # 종료일 (Combobox — DB 날짜 자동 파악)
        tk.Label(csv_body, text='종료일', bg=C['bg_panel'], fg=C['text3'],
                 font=('맑은 고딕', 9)).grid(row=0, column=4, sticky='w', pady=2)
        self._csv_end_var = tk.StringVar()
        self._csv_end_combo = ttk.Combobox(csv_body, textvariable=self._csv_end_var,
                 font=('맑은 고딕', 10), width=12, state='readonly')
        self._csv_end_combo.grid(row=0, column=5, padx=(4, 16), pady=2)

        # 파일명 미리보기
        self._csv_preview_var = tk.StringVar(value='...')
        tk.Label(csv_body, textvariable=self._csv_preview_var,
                 bg=C['bg_panel'], fg=C['green'], font=('Consolas', 10)
                 ).grid(row=0, column=6, padx=(8, 16), sticky='w')

        # 내보내기 버튼
        _dark_button(csv_body, '내보내기', self._csv_export, 'export').grid(
            row=0, column=7, padx=(0, 4), pady=2)

        self._csv_start_var.trace_add('write', self._csv_update_preview)
        self._csv_end_var.trace_add('write', self._csv_update_preview)

        # DB 이력 로드
        if self.db_path:
            try:
                self._db_records = get_recent_history(self.db_path)
            except Exception:
                self._db_records = []

        self._refresh_tree()
        root.mainloop()
        self._alive = False

    # ── Treeview 갱신 ──
    def _refresh_tree(self):
        if not self._root:
            return

        with self._lock:
            pending_copy = list(self._pending)

        # 날짜 집계
        date_set = {}
        new_dates = set()
        for _, date_str, records, _ in pending_copy:
            new_dates.add(date_str)
            date_set[date_str] = date_set.get(date_str, 0) + len(records)
        for rec in self._db_records:
            d = rec.get('date', '')
            date_set[d] = date_set.get(d, 0) + 1

        # 내보내기 상태 조회
        export_status = {}
        if self.db_path:
            try:
                export_status = get_export_status_by_date(self.db_path)
            except Exception:
                pass

        # 날짜 목록: ● 신규 / ✓ 내보내기완료 / ◐ 부분내보내기 / 빈칸 미내보내기
        self._date_listbox.delete(0, 'end')
        for d in sorted(date_set.keys(), reverse=True):
            if d in new_dates:
                prefix = '● '
            elif export_status.get(d) == 'all':
                prefix = '✓ '
            elif export_status.get(d) == 'partial':
                prefix = '◐ '
            else:
                prefix = '  '
            label = f'{prefix}{d} ({date_set[d]})'
            self._date_listbox.insert('end', label)
            idx = self._date_listbox.size() - 1
            if d in new_dates:
                self._date_listbox.itemconfig(idx, fg=C['green'])
            elif export_status.get(d) == 'all':
                self._date_listbox.itemconfig(idx, fg=C['accent'])
            elif export_status.get(d) == 'partial':
                self._date_listbox.itemconfig(idx, fg=C['amber'])

        # Treeview
        self._tree.delete(*self._tree.get_children())
        self._row_records.clear()

        def _rec_to_values(rec):
            return tuple(_tree_cell(rec, d[0]) for d in self._col_defs)

        for doc_name, date_str, records, content_hash in pending_copy:
            if self._current_filter_date and date_str != self._current_filter_date:
                continue
            for rec in records:
                # pending 레코드에 content_hash 추가 (표시용)
                rec_with_hash = dict(rec, content_hash=content_hash or '')
                iid = self._tree.insert('', 'end',
                    tags=('new', f'p:{doc_name}:{date_str}'),
                    values=_rec_to_values(rec_with_hash))
                self._row_records[iid] = rec_with_hash

        for rec in self._db_records:
            d = rec.get('date', '')
            if self._current_filter_date and d != self._current_filter_date:
                continue
            iid = self._tree.insert('', 'end', tags=(f'db:{rec.get("id","")}',),
                values=_rec_to_values(rec))
            self._row_records[iid] = rec

        total_db = len(self._db_records)
        total_new = sum(len(r) for _, _, r, _ in pending_copy)
        self._info_var.set(f'전체 이력: {total_db}건  ·  신규 파싱: {total_new}건 ({len(pending_copy)}파일)')

        # CSV 날짜 콤보박스 갱신
        self._csv_refresh_dates()

    # ── 날짜 선택 ──
    def _on_date_select(self, event):
        sel = self._date_listbox.curselection()
        if not sel:
            return
        label = self._date_listbox.get(sel[0]).strip()
        # "● 2024-05-03 (12)" or "  2024-05-03 (12)"
        date_str = label.lstrip('● ').split(' (')[0].strip()
        self._current_filter_date = date_str
        self._refresh_tree()

    def _show_all_dates(self):
        self._current_filter_date = None
        self._date_listbox.selection_clear(0, 'end')
        self._refresh_tree()

    # ── 더블클릭 편집 ──
    def _on_cell_double_click(self, event):
        item = self._tree.identify_row(event.y)
        col = self._tree.identify_column(event.x)
        if not item or not col:
            return
        col_idx = int(col.replace('#', '')) - 1
        col_name = self._col_defs[col_idx][2] if col_idx < len(self._col_defs) else ''
        col_id = self._col_defs[col_idx][0] if col_idx < len(self._col_defs) else ''
        values = list(self._tree.item(item, 'values'))
        # 팝업에는 _row_records 의 원본 값(개행 포함) 표시.
        # 폴백으로 Treeview 셀의 한 줄 표시값 사용.
        rec = self._row_records.get(item)
        if rec is not None and col_id in rec:
            current_val = str(rec.get(col_id, '') or '')
        else:
            current_val = values[col_idx] if col_idx < len(values) else ''

        edit_win = tk.Toplevel(self._root)
        edit_win.title(f'{col_name} 편집')
        edit_win.geometry('500x300')
        edit_win.configure(bg=C['bg_panel'])
        edit_win.attributes('-topmost', True)
        edit_win.grab_set()

        tk.Label(edit_win, text=f'[ {col_name} ]',
                 font=('맑은 고딕', 11, 'bold'),
                 bg=C['bg_panel'], fg=C['text_accent'],
                 padx=10, pady=8).pack(anchor='w')

        btn_frame = tk.Frame(edit_win, bg=C['bg_panel'])
        btn_frame.pack(side='bottom', fill='x', padx=10, pady=10)

        def apply_edit():
            raw = text_widget.get('1.0', 'end').strip()
            if col_id in MULTILINE_TREE_COLS:
                # 멀티라인 컬럼은 개행 보존, 빈 줄만 제거
                stored_val = '\n'.join(
                    line.rstrip() for line in raw.split('\n') if line.strip())
            else:
                stored_val = _flat(raw)
            # 원본 record 갱신 (다음 더블클릭에서도 멀티라인 유지)
            if rec is not None and col_id:
                rec[col_id] = stored_val
            # Treeview 셀은 항상 한 줄로 표시
            values[col_idx] = _tree_cell({col_id: stored_val}, col_id) if col_id else stored_val
            self._tree.item(item, values=values)
            edit_win.destroy()

        _dark_button(btn_frame, '저장', apply_edit, 'primary').pack(side='right', padx=4)
        _dark_button(btn_frame, '닫기', edit_win.destroy).pack(side='right')

        text_frame = tk.Frame(edit_win, bg=C['bg_panel'])
        text_frame.pack(fill='both', expand=True, padx=10, pady=(0, 4))

        text_widget = tk.Text(text_frame, font=('맑은 고딕', 10), wrap='word',
                              bg=C['bg_deep'], fg=C['text1'],
                              insertbackground=C['text1'],
                              highlightbackground=C['border'], highlightthickness=1,
                              relief='flat')
        text_widget.insert('1.0', current_val)
        text_widget.pack(fill='both', expand=True)

    # ── 행 삭제 ──
    def _delete_selected(self):
        selected = list(self._tree.selection())
        if not selected:
            return
        db_ids, pending_keys = [], set()
        for iid in selected:
            for tag in self._tree.item(iid, 'tags'):
                if tag.startswith('db:'):
                    try:
                        db_ids.append(int(tag.split(':', 1)[1]))
                    except ValueError:
                        pass
                elif tag.startswith('p:'):
                    parts = tag.split(':', 2)
                    if len(parts) == 3:
                        pending_keys.add((parts[1], parts[2]))
        if db_ids and self.db_path:
            delete_by_ids(self.db_path, db_ids)
            self._db_records = [r for r in self._db_records if r.get('id') not in db_ids]
        if pending_keys:
            with self._lock:
                self._pending = [(dn, ds, recs, ch) for dn, ds, recs, ch in self._pending
                                 if (dn, ds) not in pending_keys]
        self._refresh_tree()

    # ── 저장 ──
    def _save_all(self):
        with self._lock:
            pending_copy = list(self._pending)
            self._pending.clear()
        if pending_copy:
            self.on_save_all(pending_copy)
        if self.db_path:
            try:
                self._db_records = get_recent_history(self.db_path)
            except Exception:
                pass
        self._refresh_tree()

    def _csv_on_range_change(self, event=None):
        is_all = self._csv_range_var.get() == '전체 내보내기'
        state = 'disabled' if is_all else 'readonly'
        self._csv_start_combo.config(state=state)
        self._csv_end_combo.config(state=state)
        self._csv_update_preview()

    def _csv_refresh_dates(self):
        """DB에서 고유 날짜 목록을 가져와 시작일/종료일 Combobox에 설정."""
        dates = sorted(set(r.get('date', '') for r in self._db_records if r.get('date')))
        # pending 날짜도 포함
        with self._lock:
            for _, ds, _, _ in self._pending:
                if ds and ds not in dates:
                    dates.append(ds)
        dates = sorted(dates)
        self._csv_start_combo['values'] = dates
        self._csv_end_combo['values'] = dates
        if dates:
            if not self._csv_start_var.get():
                self._csv_start_var.set(dates[0])
            if not self._csv_end_var.get():
                self._csv_end_var.set(dates[-1])
        self._csv_update_preview()

    def _csv_update_preview(self, *_):
        if self._csv_range_var.get() == '전체 내보내기':
            self._csv_preview_var.set('전체.csv')
        else:
            s = self._csv_start_var.get().replace('-', '_')
            e = self._csv_end_var.get().replace('-', '_')
            if s and e:
                self._csv_preview_var.set(f'{s}_{e}.csv')
            else:
                self._csv_preview_var.set('...')

    def _csv_export(self):
        is_all = self._csv_range_var.get() == '전체 내보내기'
        sd = None if is_all else (self._csv_start_var.get().strip() or None)
        ed = None if is_all else (self._csv_end_var.get().strip() or None)
        default = '전체.csv' if is_all else self._csv_preview_var.get()
        # 기본 저장 경로: autosave 폴더
        os.makedirs(AUTOSAVE_DIR, exist_ok=True)
        path = filedialog.asksaveasfilename(parent=self._root, title='CSV 저장',
            defaultextension='.csv', filetypes=[('CSV', '*.csv')],
            initialfile=default, initialdir=AUTOSAVE_DIR)
        if not path:
            return
        count = export_csv(self.db_path, path, start_date=sd, end_date=ed)
        messagebox.showinfo('완료', f'{count}건 내보내기 완료\n경로: {path}', parent=self._root)
        # 내보내기 후 DB 이력 재로드 + 날짜 목록 갱신 (내보내기 상태 반영)
        if self.db_path:
            try:
                self._db_records = get_recent_history(self.db_path)
            except Exception:
                pass
        self._refresh_tree()

    def _on_close(self):
        self._alive = False
        if self._root:
            self._root.destroy()
            self._root = None


def ask_date_input(doc_name):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    date_str = simpledialog.askstring('날짜 입력',
        f'"{doc_name}"에서 날짜를 찾을 수 없습니다.\n날짜를 입력해 주십시오 (예: 2024-05-03):',
        parent=root)
    root.destroy()
    return date_str

# ─────────────────────────────────────────────
# Tray
# ─────────────────────────────────────────────

def create_icon_image():
    img = Image.new('RGB', (64, 64), color=(41, 98, 255))
    draw = ImageDraw.Draw(img)
    draw.rectangle([16, 10, 48, 54], outline='white', width=2)
    draw.line([22, 22, 42, 22], fill='white', width=2)
    draw.line([22, 30, 42, 30], fill='white', width=2)
    draw.line([22, 38, 36, 38], fill='white', width=2)
    return img


class TrayApp:
    def __init__(self, db_path, save_all_callback=None):
        self.db_path = db_path
        self.save_all_callback = save_all_callback or (lambda p: None)
        self.icon = None
        self.auto_save = False

    def start(self):
        menu = pystray.Menu(
            pystray.MenuItem('파싱 뷰어 열기', self._show_viewer),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('확인 후 저장', self._set_confirm,
                             checked=lambda item: not self.auto_save),
            pystray.MenuItem('즉시 저장', self._set_auto,
                             checked=lambda item: self.auto_save),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('종료', self._quit),
        )
        self.icon = pystray.Icon('설비일보 Word 파서', create_icon_image(),
                                  '설비일보 Word 파서', menu)
        self.icon.run()

    def notify(self, msg):
        if self.icon:
            self.icon.notify(msg, '설비일보 Word 파서')

    def _show_viewer(self, icon, item):
        popup = ParseResultPopup.get_or_create(
            on_save_all=self.save_all_callback, db_path=self.db_path)
        if not popup._alive:
            threading.Thread(target=popup._show, daemon=True).start()

    def _set_confirm(self, icon, item):
        self.auto_save = False

    def _set_auto(self, icon, item):
        self.auto_save = True

    def _quit(self, icon, item):
        icon.stop()

# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────

def main():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    lock_file = open(LOCK_PATH, 'w')
    try:
        msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
    except (IOError, OSError):
        print('이미 실행 중입니다.')
        lock_file.close()
        sys.exit(0)

    init_db(DB_PATH)

    def save_all(pending):
        for doc_name, date_str, records, content_hash in pending:
            delete_by_source(DB_PATH, doc_name, date_str)
            insert_records(DB_PATH, records, content_hash)
        total = sum(len(r) for _, _, r, _ in pending)
        tray.notify(f'전체 저장 완료 ({len(pending)}파일, {total}건)')

    tray = TrayApp(DB_PATH, save_all_callback=save_all)

    def on_new_parse(doc_name, date_str, records, content_hash):
        if tray.auto_save:
            delete_by_source(DB_PATH, doc_name, date_str)
            insert_records(DB_PATH, records, content_hash)
            tray.notify(f'설비일보 {date_str} 저장 완료 ({len(records)}건)')
        else:
            tray.notify(f'설비일보 {date_str} 파싱 완료 ({len(records)}건)')
            popup = ParseResultPopup.get_or_create(on_save_all=save_all, db_path=DB_PATH)
            popup.add_records(doc_name, date_str, records, content_hash)

    def on_duplicate_same(doc_name, date_str):
        tray.notify(f'이미 파싱된 파일입니다 ({date_str})')

    def on_duplicate_changed(doc_name, date_str, records, content_hash):
        tray.notify(f'내용 변경 감지 ({date_str})')
        popup = ParseResultPopup.get_or_create(on_save_all=save_all, db_path=DB_PATH)
        popup.add_records(doc_name, date_str, records, content_hash)

    def on_no_table(doc_name):
        tray.notify(f'파싱 대상 아님: {doc_name}')

    def on_date_missing(doc_name):
        return ask_date_input(doc_name)

    watcher = WordWatcher(
        db_path=DB_PATH, on_new_parse=on_new_parse,
        on_duplicate_same=on_duplicate_same,
        on_duplicate_changed=on_duplicate_changed,
        on_date_missing=on_date_missing, on_no_table=on_no_table)
    watcher.start()
    tray.start()
    watcher.stop()

    try:
        msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
    except Exception:
        pass
    lock_file.close()
    try:
        os.remove(LOCK_PATH)
    except Exception:
        pass


if __name__ == '__main__':
    main()
