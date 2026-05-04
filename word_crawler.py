"""
설비일보 Word 크롤러 (단일 파일)
- 열린 Word 문서에서 설비일보 표를 자동 감지/파싱 (동적 헤더)
- [현상][원인][조치] 태그 추출 + 2+ 개행으로 항목 분리
- SQLite 적재 + CSV 내보내기 → SQream 이관
- 시스템 트레이 + tkinter 다크 테마 팝업 UI
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

TAG_PATTERN = re.compile(r'\[(현상|원인|조치)\]\s*([^\r\n\[\]]+)')

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
            phenomenon       TEXT,
            cause            TEXT,
            action           TEXT,
            raw_text         TEXT,
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
             title, phenomenon, cause, action, raw_text, raw_cell, header4, val4, content_hash)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            rec['date'], rec['source_file'], rec['row_num'],
            rec['header1'], rec['val1'], rec['content_col_name'],
            rec.get('title', ''), rec.get('phenomenon', ''), rec.get('cause', ''), rec.get('action', ''),
            rec.get('raw_text', ''), rec['raw_cell'],
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


def parse_item_block(text):
    """한 블록(2+ 개행으로 분리된 단위)에서 title + [현상][원인][조치] 태그 추출.
    [현상] 앞의 텍스트 → title. 태그가 없으면 raw_text에 전체 텍스트."""
    text = _strip_numbering(text)
    parsed = {'title': '', 'phenomenon': '', 'cause': '', 'action': '', 'raw_text': ''}

    # [현상] 앞의 텍스트를 title로 추출
    first_tag = re.search(r'\[(현상|원인|조치)\]', text)
    if first_tag:
        before = text[:first_tag.start()].strip()
        if before:
            parsed['title'] = before

    for match in TAG_PATTERN.finditer(text):
        tag, content = match.group(1), match.group(2).strip()
        if tag == '현상':
            parsed['phenomenon'] = content
        elif tag == '원인':
            parsed['cause'] = content
        elif tag == '조치':
            parsed['action'] = content

    if not parsed['phenomenon'] and not parsed['cause'] and not parsed['action']:
        parsed['raw_text'] = text.strip()
    else:
        leftover = re.sub(r'\[(현상|원인|조치)\]\s*[^\r\n\[\]]+', '', text).strip()
        leftover = '\n'.join(line.strip() for line in leftover.split('\n') if line.strip())
        # title로 이미 추출한 부분 제거
        if parsed['title'] and leftover.startswith(parsed['title']):
            leftover = leftover[len(parsed['title']):].strip()
        parsed['raw_text'] = leftover

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
                    'phenomenon': parsed['phenomenon'],
                    'cause': parsed['cause'],
                    'action': parsed['action'],
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
        row_counts = []
        for t in range(1, doc.Tables.Count + 1):
            try:
                row_counts.append(doc.Tables(t).Rows.Count)
            except Exception:
                row_counts.append(0)
        table_idx = find_main_table_index(row_counts)
        if table_idx is None:
            return None
        table = doc.Tables(table_idx + 1)
        headers = []
        row1 = table.Rows(1)
        for c in range(1, row1.Cells.Count + 1):
            headers.append(clean_cell_text(row1.Cells(c).Range.Text))
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
        rows_data = []
        for r in range(2, table.Rows.Count + 1):
            row = table.Rows(r)
            cells = []
            for c in range(1, row.Cells.Count + 1):
                cells.append(clean_cell_text(row.Cells(c).Range.Text))
            rows_data.append(cells)
        records = parse_table_data(headers, rows_data, date_str, doc.Name)
        return (records, date_str) if records else None

    def reset_seen(self):
        self._seen_docs.clear()

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────

def _flat(text):
    if not text:
        return ''
    return ' '.join(str(text).replace('\r', ' ').replace('\n', ' ').split())


def _apply_dark_theme(root):
    root.configure(bg=C['bg_deep'])
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('Treeview',
                    background=C['bg_panel'], foreground=C['text1'],
                    fieldbackground=C['bg_panel'], borderwidth=0,
                    font=('맑은 고딕', 9), rowheight=24)
    style.configure('Treeview.Heading',
                    background=C['bg_surface'], foreground=C['text_accent'],
                    font=('맑은 고딕', 9, 'bold'), borderwidth=0)
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

        cols = ('date', 'val1', 'col_name', 'title', 'phenomenon', 'cause', 'action', 'raw_text', 'val4')
        col_cfg = {
            'date':       ('날짜', 85),
            'val1':       ('구분', 60),
            'col_name':   ('영역', 60),
            'title':      ('제목', 150),
            'phenomenon': ('현상', 180),
            'cause':      ('원인', 180),
            'action':     ('조치', 180),
            'raw_text':   ('기타', 180),
            'val4':       ('비고', 100),
        }

        tree_frame = _dark_frame(right)
        tree_frame.pack(fill='both', expand=True)

        sy = ttk.Scrollbar(tree_frame, orient='vertical')
        sx = ttk.Scrollbar(tree_frame, orient='horizontal')
        self._tree = ttk.Treeview(tree_frame, columns=cols, show='headings',
                                   yscrollcommand=sy.set, xscrollcommand=sx.set,
                                   selectmode='extended')
        sy.config(command=self._tree.yview)
        sx.config(command=self._tree.xview)

        for col in cols:
            label, w = col_cfg[col]
            self._tree.heading(col, text=label)
            self._tree.column(col, width=w, minwidth=40, stretch=True)

        self._tree.tag_configure('new', background=C['green_row'])
        self._tree.bind('<Double-1>', self._on_cell_double_click)

        sy.pack(side='right', fill='y')
        sx.pack(side='bottom', fill='x')
        self._tree.pack(fill='both', expand=True)

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

        # 시작일
        tk.Label(csv_body, text='시작일', bg=C['bg_panel'], fg=C['text3'],
                 font=('맑은 고딕', 9)).grid(row=0, column=2, sticky='w', pady=2)
        self._csv_start_var = tk.StringVar()
        self._csv_start_entry = tk.Entry(csv_body, textvariable=self._csv_start_var,
                 font=('맑은 고딕', 10), width=12,
                 bg=C['bg_deep'], fg=C['text1'], insertbackground=C['text1'],
                 highlightbackground=C['border'], highlightthickness=1, relief='flat')
        self._csv_start_entry.grid(row=0, column=3, padx=(4, 16), pady=2)

        # 종료일
        tk.Label(csv_body, text='종료일', bg=C['bg_panel'], fg=C['text3'],
                 font=('맑은 고딕', 9)).grid(row=0, column=4, sticky='w', pady=2)
        self._csv_end_var = tk.StringVar()
        self._csv_end_entry = tk.Entry(csv_body, textvariable=self._csv_end_var,
                 font=('맑은 고딕', 10), width=12,
                 bg=C['bg_deep'], fg=C['text1'], insertbackground=C['text1'],
                 highlightbackground=C['border'], highlightthickness=1, relief='flat')
        self._csv_end_entry.grid(row=0, column=5, padx=(4, 16), pady=2)

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

        # 날짜 목록
        self._date_listbox.delete(0, 'end')
        for d in sorted(date_set.keys(), reverse=True):
            label = f'{"● " if d in new_dates else "  "}{d} ({date_set[d]})'
            self._date_listbox.insert('end', label)
            if d in new_dates:
                idx = self._date_listbox.size() - 1
                self._date_listbox.itemconfig(idx, fg=C['green'])

        # Treeview
        self._tree.delete(*self._tree.get_children())

        for doc_name, date_str, records, _ in pending_copy:
            if self._current_filter_date and date_str != self._current_filter_date:
                continue
            for rec in records:
                self._tree.insert('', 'end', tags=('new', f'p:{doc_name}:{date_str}'),
                    values=(rec.get('date',''), _flat(rec.get('val1','')),
                            _flat(rec.get('content_col_name','')),
                            _flat(rec.get('title','')),
                            _flat(rec.get('phenomenon','')),
                            _flat(rec.get('cause','')),
                            _flat(rec.get('action','')),
                            _flat(rec.get('raw_text','')),
                            _flat(rec.get('val4',''))))

        for rec in self._db_records:
            d = rec.get('date', '')
            if self._current_filter_date and d != self._current_filter_date:
                continue
            self._tree.insert('', 'end', tags=(f'db:{rec.get("id","")}',),
                values=(d, _flat(rec.get('val1','')),
                        _flat(rec.get('content_col_name','')),
                        _flat(rec.get('title','')),
                        _flat(rec.get('phenomenon','')),
                        _flat(rec.get('cause','')),
                        _flat(rec.get('action','')),
                        _flat(rec.get('raw_text','')),
                        _flat(rec.get('val4',''))))

        total_db = len(self._db_records)
        total_new = sum(len(r) for _, _, r, _ in pending_copy)
        self._info_var.set(f'전체 이력: {total_db}건  ·  신규 파싱: {total_new}건 ({len(pending_copy)}파일)')

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
        col_names = ['날짜', '구분', '영역', '제목', '현상', '원인', '조치', '기타', '비고']
        col_name = col_names[col_idx] if col_idx < len(col_names) else ''
        values = list(self._tree.item(item, 'values'))
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
            new_val = _flat(text_widget.get('1.0', 'end').strip())
            values[col_idx] = new_val
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
        selected = self._tree.selection()
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
        state = 'disabled' if is_all else 'normal'
        self._csv_start_entry.config(state=state)
        self._csv_end_entry.config(state=state)
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
        path = filedialog.asksaveasfilename(parent=self._root, title='CSV 저장',
            defaultextension='.csv', filetypes=[('CSV', '*.csv')], initialfile=default)
        if not path:
            return
        count = export_csv(self.db_path, path, start_date=sd, end_date=ed)
        messagebox.showinfo('완료', f'{count}건 내보내기 완료', parent=self._root)

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
