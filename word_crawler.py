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
from tkinter import ttk, simpledialog, filedialog, messagebox

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

        row_counts = []
        for t in range(1, table_count + 1):
            try:
                row_counts.append(doc.Tables(t).Rows.Count)
            except Exception:
                row_counts.append(0)

        table_idx = find_main_table_index(row_counts)
        if table_idx is None:
            return None

        table = doc.Tables(table_idx + 1)

        # 헤더 동적 읽기
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

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────

def _flat(text):
    """줄바꿈을 공백으로 치환하여 Treeview 한 줄 표시."""
    if text is None:
        return ''
    return str(text).replace('\n', ' ').replace('\r', ' ')


class ParseResultPopup:
    _instance = None
    _lock = threading.Lock()

    @classmethod
    def get_or_create(cls, on_save_all=None, db_path=None):
        with cls._lock:
            if cls._instance is None or not cls._instance._alive:
                cls._instance = cls(on_save_all=on_save_all, db_path=db_path)
            return cls._instance

    def __init__(self, on_save_all=None, db_path=None):
        self.on_save_all = on_save_all or (lambda p: None)
        self.db_path = db_path
        self._alive = False
        self._pending = []  # list of (doc_name, date_str, records, content_hash)
        self._lock = threading.Lock()
        self._root = None

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
        root.title('설비일보 파싱 결과')
        root.geometry('1100x600')
        root.resizable(True, True)
        root.protocol('WM_DELETE_WINDOW', self._on_close)

        # 상단 정보 라벨
        self._info_var = tk.StringVar(value='전체 이력 0건 | 신규 파싱 0건')
        info_label = tk.Label(root, textvariable=self._info_var,
                              font=('맑은 고딕', 10), anchor='w', padx=8)
        info_label.pack(side='top', fill='x', pady=(6, 0))

        # 메인 영역 (좌측 날짜 패널 + 우측 Treeview)
        main_frame = tk.Frame(root)
        main_frame.pack(side='top', fill='both', expand=True, padx=6, pady=4)

        # 좌측 날짜 패널
        left_frame = tk.Frame(main_frame, width=160)
        left_frame.pack(side='left', fill='y', padx=(0, 4))
        left_frame.pack_propagate(False)

        tk.Label(left_frame, text='날짜 목록', font=('맑은 고딕', 9, 'bold')).pack(
            side='top', anchor='w', padx=4, pady=(4, 2))

        self._date_listbox = tk.Listbox(left_frame, selectmode='single',
                                        font=('맑은 고딕', 9))
        self._date_listbox.pack(side='top', fill='both', expand=True, padx=2)
        self._date_listbox.bind('<<ListboxSelect>>', self._on_date_select)

        all_btn = tk.Button(left_frame, text='전체 보기',
                            command=self._show_all_dates)
        all_btn.pack(side='top', fill='x', padx=2, pady=4)

        # 우측 Treeview
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side='left', fill='both', expand=True)

        columns = ('date', 'val1', 'content_col_name', 'item_text', 'raw_cell', 'val4')
        col_labels = {
            'date': '날짜',
            'val1': '구분',
            'content_col_name': '영역',
            'item_text': '항목',
            'raw_cell': '원문',
            'val4': '비고',
        }
        col_widths = {
            'date': 90,
            'val1': 80,
            'content_col_name': 80,
            'item_text': 260,
            'raw_cell': 260,
            'val4': 120,
        }

        tree_scroll_y = tk.Scrollbar(right_frame, orient='vertical')
        tree_scroll_x = tk.Scrollbar(right_frame, orient='horizontal')

        self._tree = ttk.Treeview(
            right_frame,
            columns=columns,
            show='headings',
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set,
            selectmode='extended',
        )
        tree_scroll_y.config(command=self._tree.yview)
        tree_scroll_x.config(command=self._tree.xview)

        for col in columns:
            self._tree.heading(col, text=col_labels[col])
            self._tree.column(col, width=col_widths[col], minwidth=40, stretch=True)

        self._tree.tag_configure('new', background='#e8f5e9')

        tree_scroll_y.pack(side='right', fill='y')
        tree_scroll_x.pack(side='bottom', fill='x')
        self._tree.pack(side='left', fill='both', expand=True)

        # 하단 버튼 영역
        btn_frame = tk.Frame(root)
        btn_frame.pack(side='bottom', fill='x', padx=6, pady=6)

        del_btn = tk.Button(btn_frame, text='선택 행 삭제',
                            command=self._delete_selected)
        del_btn.pack(side='left')

        csv_btn = tk.Button(btn_frame, text='CSV 내보내기',
                            command=self._open_csv_dialog)
        csv_btn.pack(side='left', padx=(8, 0))

        save_btn = tk.Button(btn_frame, text='저장', width=10,
                             command=self._save_all)
        save_btn.pack(side='right')

        skip_btn = tk.Button(btn_frame, text='스킵', width=10,
                             command=self._on_close)
        skip_btn.pack(side='right', padx=(0, 4))

        # DB 이력 로드
        self._db_records = []
        if self.db_path:
            try:
                self._db_records = get_recent_history(self.db_path)
            except Exception:
                self._db_records = []

        self._current_filter_date = None
        self._refresh_tree()
        root.mainloop()
        self._alive = False

    def _refresh_tree(self):
        if self._root is None:
            return

        # 날짜 목록 갱신
        date_set = {}
        with self._lock:
            pending_copy = list(self._pending)

        new_dates = set()
        for doc_name, date_str, records, content_hash in pending_copy:
            new_dates.add(date_str)
            date_set[date_str] = date_set.get(date_str, 0) + len(records)

        for rec in self._db_records:
            d = rec.get('date', '')
            if d not in date_set:
                date_set[d] = 0
            date_set[d] += 1

        self._date_listbox.delete(0, 'end')
        sorted_dates = sorted(date_set.keys(), reverse=True)
        for d in sorted_dates:
            label = f'{d} ({date_set[d]})'
            self._date_listbox.insert('end', label)
            if d in new_dates:
                idx = self._date_listbox.size() - 1
                self._date_listbox.itemconfig(idx, {'bg': '#c8e6c9', 'fg': 'black'})

        # Treeview 갱신
        self._tree.delete(*self._tree.get_children())

        # pending 행 표시
        for doc_name, date_str, records, content_hash in pending_copy:
            if self._current_filter_date and date_str != self._current_filter_date:
                continue
            for rec in records:
                rec_id = rec.get('id', '')
                iid = self._tree.insert('', 'end', tags=('new',), values=(
                    _flat(rec.get('date', '')),
                    _flat(rec.get('val1', '')),
                    _flat(rec.get('content_col_name', '')),
                    _flat(rec.get('item_text', '')),
                    _flat(rec.get('raw_cell', '')),
                    _flat(rec.get('val4', '')),
                ))
                self._tree.item(iid, tags=('new', f'pending:{doc_name}:{date_str}'))

        # DB 이력 행 표시
        for rec in self._db_records:
            d = rec.get('date', '')
            if self._current_filter_date and d != self._current_filter_date:
                continue
            rec_id = rec.get('id', '')
            iid = self._tree.insert('', 'end', values=(
                _flat(d),
                _flat(rec.get('val1', '')),
                _flat(rec.get('content_col_name', '')),
                _flat(rec.get('item_text', '')),
                _flat(rec.get('raw_cell', '')),
                _flat(rec.get('val4', '')),
            ), tags=(f'dbid:{rec_id}',))

        # 정보 라벨 갱신
        total_history = len(self._db_records)
        total_new = sum(len(r) for _, _, r, _ in pending_copy)
        self._info_var.set(f'전체 이력 {total_history}건 | 신규 파싱 {total_new}건')

    def _on_date_select(self, event):
        sel = self._date_listbox.curselection()
        if not sel:
            return
        label = self._date_listbox.get(sel[0])
        date_str = label.split(' ')[0]
        self._current_filter_date = date_str
        self._refresh_tree()

    def _show_all_dates(self):
        self._current_filter_date = None
        self._date_listbox.selection_clear(0, 'end')
        self._refresh_tree()

    def _delete_selected(self):
        selected = self._tree.selection()
        if not selected:
            return

        db_ids_to_delete = []
        pending_keys_to_remove = set()

        for iid in selected:
            tags = self._tree.item(iid, 'tags')
            for tag in tags:
                if tag.startswith('dbid:'):
                    try:
                        db_ids_to_delete.append(int(tag.split(':', 1)[1]))
                    except ValueError:
                        pass
                elif tag.startswith('pending:'):
                    parts = tag.split(':', 2)
                    if len(parts) == 3:
                        pending_keys_to_remove.add((parts[1], parts[2]))

        # DB 행 삭제
        if db_ids_to_delete and self.db_path:
            delete_by_ids(self.db_path, db_ids_to_delete)
            self._db_records = [r for r in self._db_records
                                 if r.get('id') not in db_ids_to_delete]

        # pending 행 제거
        if pending_keys_to_remove:
            with self._lock:
                self._pending = [
                    (dn, ds, recs, ch) for dn, ds, recs, ch in self._pending
                    if (dn, ds) not in pending_keys_to_remove
                ]

        self._refresh_tree()

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

    def _open_csv_dialog(self):
        if self._root:
            CsvExportDialog(self._root, self.db_path)

    def _on_close(self):
        self._alive = False
        if self._root:
            self._root.destroy()
            self._root = None


class CsvExportDialog:
    def __init__(self, parent, db_path):
        self.db_path = db_path
        self.win = tk.Toplevel(parent)
        self.win.title('CSV 내보내기')
        self.win.geometry('380x220')
        self.win.resizable(False, False)
        self.win.grab_set()

        # 범위 선택
        range_frame = tk.Frame(self.win)
        range_frame.pack(fill='x', padx=16, pady=(16, 4))
        tk.Label(range_frame, text='범위:').pack(side='left')
        self._range_var = tk.StringVar(value='날짜 범위 지정')
        self._range_combo = ttk.Combobox(
            range_frame,
            textvariable=self._range_var,
            values=['날짜 범위 지정', '전체 내보내기'],
            state='readonly',
            width=16,
        )
        self._range_combo.pack(side='left', padx=(8, 0))
        self._range_combo.bind('<<ComboboxSelected>>', self._on_range_change)

        # 시작일/종료일
        date_frame = tk.Frame(self.win)
        date_frame.pack(fill='x', padx=16, pady=4)

        tk.Label(date_frame, text='시작일:').grid(row=0, column=0, sticky='w', pady=2)
        self._start_var = tk.StringVar()
        self._start_entry = tk.Entry(date_frame, textvariable=self._start_var, width=14)
        self._start_entry.grid(row=0, column=1, padx=(8, 0), sticky='w')

        tk.Label(date_frame, text='종료일:').grid(row=1, column=0, sticky='w', pady=2)
        self._end_var = tk.StringVar()
        self._end_entry = tk.Entry(date_frame, textvariable=self._end_var, width=14)
        self._end_entry.grid(row=1, column=1, padx=(8, 0), sticky='w')

        # 파일명 미리보기
        self._preview_var = tk.StringVar(value='')
        preview_label = tk.Label(self.win, textvariable=self._preview_var,
                                  font=('맑은 고딕', 9), fg='gray')
        preview_label.pack(fill='x', padx=16, pady=(4, 0))

        self._start_var.trace_add('write', self._update_preview)
        self._end_var.trace_add('write', self._update_preview)
        self._range_var.trace_add('write', self._update_preview)

        # 버튼
        btn_frame = tk.Frame(self.win)
        btn_frame.pack(side='bottom', fill='x', padx=16, pady=12)

        tk.Button(btn_frame, text='취소', width=10,
                  command=self.win.destroy).pack(side='right')
        tk.Button(btn_frame, text='내보내기', width=10,
                  command=self._export).pack(side='right', padx=(0, 4))

        self._update_preview()

    def _on_range_change(self, event=None):
        if self._range_var.get() == '전체 내보내기':
            self._start_entry.config(state='disabled')
            self._end_entry.config(state='disabled')
        else:
            self._start_entry.config(state='normal')
            self._end_entry.config(state='normal')
        self._update_preview()

    def _update_preview(self, *args):
        if self._range_var.get() == '전체 내보내기':
            self._preview_var.set('파일명: 전체.csv')
        else:
            s = self._start_var.get().replace('-', '_')
            e = self._end_var.get().replace('-', '_')
            if s and e:
                self._preview_var.set(f'파일명: {s}_{e}.csv')
            elif s:
                self._preview_var.set(f'파일명: {s}_~.csv')
            elif e:
                self._preview_var.set(f'파일명: ~_{e}.csv')
            else:
                self._preview_var.set('파일명: (미리보기)')

    def _export(self):
        is_all = self._range_var.get() == '전체 내보내기'
        start_date = None if is_all else (self._start_var.get().strip() or None)
        end_date = None if is_all else (self._end_var.get().strip() or None)

        if is_all:
            default_name = '전체.csv'
        else:
            s = (start_date or '').replace('-', '_')
            e = (end_date or '').replace('-', '_')
            default_name = f'{s}_{e}.csv' if (s or e) else 'export.csv'

        output_path = filedialog.asksaveasfilename(
            parent=self.win,
            title='CSV 저장',
            defaultextension='.csv',
            filetypes=[('CSV 파일', '*.csv'), ('모든 파일', '*.*')],
            initialfile=default_name,
        )
        if not output_path:
            return

        count = export_csv(self.db_path, output_path,
                           start_date=start_date, end_date=end_date)
        tk.messagebox.showinfo('완료', f'{count}건 내보내기 완료', parent=self.win)
        self.win.destroy()


def ask_date_input(doc_name):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    date_str = simpledialog.askstring(
        '날짜 입력',
        f'"{doc_name}"에서 날짜를 찾을 수 없습니다.\n날짜를 입력해 주십시오 (예: 2024-05-03):',
        parent=root
    )
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
        self.auto_save = False  # False = 확인 후 저장 (기본)

    def start(self):
        menu = pystray.Menu(
            pystray.MenuItem('파싱 뷰어 열기', self._show_viewer),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('확인 후 저장', self._set_confirm_mode,
                             checked=lambda item: not self.auto_save),
            pystray.MenuItem('즉시 저장', self._set_auto_mode,
                             checked=lambda item: self.auto_save),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('종료', self._quit),
        )
        self.icon = pystray.Icon(
            '설비일보 Word 파서', create_icon_image(),
            '설비일보 Word 파서', menu
        )
        self.icon.run()

    def notify(self, message):
        if self.icon:
            self.icon.notify(message, '설비일보 Word 파서')

    def _show_viewer(self, icon, item):
        popup = ParseResultPopup.get_or_create(
            on_save_all=self.save_all_callback, db_path=self.db_path
        )
        if not popup._alive:
            threading.Thread(target=popup._show, daemon=True).start()

    def _set_confirm_mode(self, icon, item):
        self.auto_save = False

    def _set_auto_mode(self, icon, item):
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
            # 즉시 저장 모드
            delete_by_source(DB_PATH, doc_name, date_str)
            insert_records(DB_PATH, records, content_hash)
            tray.notify(f'설비일보 {date_str} 저장 완료 ({len(records)}건)')
        else:
            # 확인 후 저장 모드
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
        db_path=DB_PATH,
        on_new_parse=on_new_parse,
        on_duplicate_same=on_duplicate_same,
        on_duplicate_changed=on_duplicate_changed,
        on_date_missing=on_date_missing,
        on_no_table=on_no_table,
    )
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
