"""
설비일보 Word 크롤러 (단일 파일)
- 열린 Word 문서에서 설비일보 표를 자동 감지/파싱
- SQLite에 적재
- 시스템 트레이 + tkinter 팝업 UI
"""
import os
import sys
import re
import sqlite3
import hashlib
import threading
import msvcrt

import pythoncom
import win32com.client
import pystray
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import ttk, simpledialog

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

EQUIP_KEYWORDS_COMPOUND = [
    '송풍팬', '배기팬', '급기팬', '환기팬',
    '순환펌프', '급수펌프', '배수펌프',
]

EQUIP_KEYWORDS_SINGLE = [
    'AHU', 'EHP', 'FCU', 'PAC', 'VAV', 'HVAC',
    '보일러', '펌프', '팬', '냉각탑', '냉동기',
    '댐퍼', '덕트', '컴프레서', '밸브', '인버터',
    '액추에이터', '팽창탱크', '집수정',
]

TAG_PATTERN = re.compile(r'\[(현상|원인|조치)\]\s*([^\r\n\[\]]+)')
ITEM_NUM_PATTERN = re.compile(r'(?=\d+\)\s)')

DATE_PATTERNS_TEXT = [
    re.compile(r'(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})'),
    re.compile(r'(\d{2})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})'),
]

DATE_PATTERNS_FILENAME = [
    re.compile(r'(\d{4})(\d{2})(\d{2})'),
    re.compile(r'(\d{2})(\d{2})(\d{2})'),
]

TABLE_HEADER_KEYWORDS = ['구분', 'A동', 'B동']

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
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            date        TEXT NOT NULL,
            shift       TEXT NOT NULL,
            building    TEXT NOT NULL,
            equipment   TEXT,
            phenomenon  TEXT,
            cause       TEXT,
            action      TEXT,
            raw_text    TEXT NOT NULL,
            remark      TEXT,
            source_file TEXT NOT NULL,
            content_hash TEXT,
            created_at  TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_fd_date ON facility_daily(date)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_fd_equipment ON facility_daily(equipment)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_fd_source ON facility_daily(source_file)')
    conn.commit()
    conn.close()


def insert_records(db_path, records, content_hash=None):
    conn = sqlite3.connect(db_path)
    for rec in records:
        conn.execute('''
            INSERT INTO facility_daily
            (date, shift, building, equipment, phenomenon, cause, action, raw_text, remark, source_file, content_hash)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            rec['date'], rec['shift'], rec['building'], rec['equipment'],
            rec['phenomenon'], rec['cause'], rec['action'],
            rec['raw_text'], rec['remark'], rec['source_file'], content_hash
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
        f"{r['shift']}:{r['building']}:{r['raw_text']}" for r in records
    )
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()[:16]

# ─────────────────────────────────────────────
# Parser
# ─────────────────────────────────────────────

def extract_equipment(text):
    found = []
    lower = text.lower()
    for kw in EQUIP_KEYWORDS_COMPOUND:
        if kw.lower() in lower:
            found.append(kw)
    for kw in EQUIP_KEYWORDS_SINGLE:
        if kw.lower() in lower:
            if not any(kw.lower() in c.lower() for c in found):
                found.append(kw)
    return found if found else ['기타']


def parse_cell(cell_text):
    text = cell_text.strip()
    if not text:
        return []

    records = []
    numbered_parts = ITEM_NUM_PATTERN.split(text)
    numbered_parts = [p.strip() for p in numbered_parts if p.strip()]
    has_numbered = any(re.match(r'\d+\)\s', p) for p in numbered_parts)

    if has_numbered:
        remaining_parts = []
        for part in numbered_parts:
            if re.match(r'\d+\)\s', part):
                dash_split = re.split(r'\n(?=-\s)', part, maxsplit=1)
                records.append(_parse_single_item(dash_split[0]))
                if len(dash_split) > 1:
                    remaining_parts.append(dash_split[1])
            else:
                remaining_parts.append(part)
        remaining = '\n'.join(remaining_parts).strip()
        if remaining and re.match(r'^-\s+', remaining):
            records.append(_parse_dash_event(remaining))
    elif text.startswith('-'):
        records.append(_parse_dash_event(text))
    else:
        records.append(_parse_single_item(text))

    return records


def _parse_single_item(text):
    parsed = {'phenomenon': '', 'cause': '', 'action': ''}
    for match in TAG_PATTERN.finditer(text):
        tag, content = match.group(1), match.group(2).strip()
        if tag == '현상':
            parsed['phenomenon'] = content
        elif tag == '원인':
            parsed['cause'] = content
        elif tag == '조치':
            parsed['action'] = content
    raw = re.sub(r'^\d+\)\s*', '', text).strip()
    raw = re.sub(r'\[(현상|원인|조치)\]\s*', '', raw).strip()
    equips = extract_equipment(text)
    return {
        'equipment': ', '.join(equips),
        'phenomenon': parsed['phenomenon'],
        'cause': parsed['cause'],
        'action': parsed['action'],
        'raw_text': raw,
    }


def _parse_dash_event(text):
    lines = text.strip().split('\n')
    combined = ' '.join(line.strip().lstrip('-').strip() for line in lines if line.strip())
    equips = extract_equipment(combined)
    return {
        'equipment': ', '.join(equips),
        'phenomenon': '',
        'cause': '',
        'action': '',
        'raw_text': combined,
    }


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


def identify_target_table_index(headers_list):
    for i, headers in enumerate(headers_list):
        header_text = ' '.join(headers)
        if all(kw in header_text for kw in TABLE_HEADER_KEYWORDS):
            return i
    return None

# ─────────────────────────────────────────────
# Word Watcher
# ─────────────────────────────────────────────

class WordWatcher:
    def __init__(self, db_path, on_new_parse, on_duplicate_same, on_duplicate_changed, on_date_missing):
        self.db_path = db_path
        self.on_new_parse = on_new_parse
        self.on_duplicate_same = on_duplicate_same
        self.on_duplicate_changed = on_duplicate_changed
        self.on_date_missing = on_date_missing
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
            doc_name = doc.Name
            doc_full = doc.FullName
            if doc_full in self._seen_docs:
                continue

            result = self._parse_document(doc)
            if result is None:
                continue

            self._seen_docs.add(doc_full)
            records, date_str = result
            content_hash = compute_hash(records)
            dup_status = check_duplicate(self.db_path, doc_name, date_str, content_hash)

            if dup_status == 'new':
                self.on_new_parse(doc_name, date_str, records, content_hash)
            elif dup_status == 'same':
                self.on_duplicate_same(doc_name, date_str)
            else:
                self.on_duplicate_changed(doc_name, date_str, records, content_hash)

    def _parse_document(self, doc):
        headers_list = []
        for t in range(1, doc.Tables.Count + 1):
            table_obj = doc.Tables(t)
            row1 = table_obj.Rows(1)
            headers = []
            for c in range(1, row1.Cells.Count + 1):
                h_raw = row1.Cells(c).Range.Text
                h_text = ''.join(ch for ch in h_raw if ord(ch) >= 32).strip()
                headers.append(h_text)
            headers_list.append(headers)

        table_idx = identify_target_table_index(headers_list)
        if table_idx is None:
            return None

        table = doc.Tables(table_idx + 1)

        date_str = None
        table_range = table.Range
        doc_text_before_table = doc.Range(0, table_range.Start).Text
        date_str = extract_date_from_text(doc_text_before_table)
        if date_str is None:
            date_str = extract_date_from_filename(doc.Name)
        if date_str is None:
            date_str = self.on_date_missing(doc.Name)
            if date_str is None:
                return None

        records = []
        num_rows = table.Rows.Count
        for r in range(2, num_rows + 1):
            row = table.Rows(r)
            cells = []
            for c in range(1, row.Cells.Count + 1):
                raw = row.Cells(c).Range.Text
                cell_text = raw.replace('\r\x07', '').replace('\x07', '').replace('\x0b', '\n').replace('\r', '\n')
                cell_text = ''.join(ch if ch == '\n' or (ord(ch) >= 32) else '' for ch in cell_text).strip()
                cells.append(cell_text)

            shift = cells[0] if len(cells) > 0 else ''
            remark = cells[3] if len(cells) > 3 else ''

            for col_idx, building in [(1, 'A동'), (2, 'B동')]:
                if col_idx >= len(cells):
                    continue
                parsed_items = parse_cell(cells[col_idx])
                for item in parsed_items:
                    records.append({
                        'date': date_str,
                        'shift': shift,
                        'building': building,
                        'equipment': item['equipment'],
                        'phenomenon': item['phenomenon'],
                        'cause': item['cause'],
                        'action': item['action'],
                        'raw_text': item['raw_text'],
                        'remark': remark,
                        'source_file': doc.Name,
                    })

        return (records, date_str) if records else None

    def reset_seen(self):
        self._seen_docs.clear()

# ─────────────────────────────────────────────
# UI (tkinter)
# ─────────────────────────────────────────────

def _flat(text):
    if not text:
        return ''
    return ' '.join(text.replace('\r', ' ').replace('\n', ' ').split())


class ParseResultPopup:
    _instance = None
    _lock = threading.Lock()

    @classmethod
    def get_or_create(cls, on_save_all, db_path=None):
        with cls._lock:
            if cls._instance is not None and cls._instance._alive:
                return cls._instance
            inst = cls(on_save_all, db_path)
            cls._instance = inst
            return inst

    def __init__(self, on_save_all, db_path=None):
        self.on_save_all = on_save_all
        self.db_path = db_path
        self._pending = []
        self._all_records = {}
        self._pending_dates = set()
        self._alive = False
        self.root = None
        self.tree = None
        self.date_listbox = None
        self.info_label = None

    def add_records(self, doc_name, date_str, records, content_hash):
        self._pending.append((doc_name, date_str, records, content_hash))
        self._pending_dates.add(date_str)
        if date_str not in self._all_records:
            self._all_records[date_str] = []
        self._all_records[date_str].extend(records)

        if self._alive and self.root:
            self.root.after(0, self._refresh_date_list)
            self.root.after(0, lambda: self._select_date(date_str))
            self.root.after(0, self._update_info)
        else:
            threading.Thread(target=self._show, daemon=True).start()

    def _show(self):
        self._alive = True
        self.root = tk.Tk()
        self.root.title('설비일보 파서 (Word)')
        self.root.geometry('1400x600')
        self.root.attributes('-topmost', True)
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

        info_frame = tk.Frame(self.root, padx=10, pady=5)
        info_frame.pack(fill='x')
        self.info_label = tk.Label(info_frame, text='', anchor='w', font=('맑은 고딕', 10))
        self.info_label.pack(fill='x')

        main_frame = tk.Frame(self.root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)

        left_frame = tk.Frame(main_frame, width=140)
        left_frame.pack(side='left', fill='y', padx=(0, 5))
        left_frame.pack_propagate(False)
        tk.Label(left_frame, text='날짜 목록', font=('맑은 고딕', 10, 'bold'), anchor='w').pack(fill='x', pady=(0, 5))

        listbox_frame = tk.Frame(left_frame)
        listbox_frame.pack(fill='both', expand=True)
        self.date_listbox = tk.Listbox(listbox_frame, font=('맑은 고딕', 10), selectmode='single')
        date_scrollbar = ttk.Scrollbar(listbox_frame, orient='vertical', command=self.date_listbox.yview)
        self.date_listbox.configure(yscrollcommand=date_scrollbar.set)
        self.date_listbox.pack(side='left', fill='both', expand=True)
        date_scrollbar.pack(side='right', fill='y')
        self.date_listbox.bind('<<ListboxSelect>>', self._on_date_select)
        tk.Button(left_frame, text='전체 보기', width=14, command=self._show_all, font=('맑은 고딕', 9)).pack(pady=(5, 0))

        right_frame = tk.Frame(main_frame)
        right_frame.pack(side='left', fill='both', expand=True)

        columns = ('date', 'shift', 'building', 'equipment', 'phenomenon', 'cause', 'action', 'raw_text')
        self.tree = ttk.Treeview(right_frame, columns=columns, show='headings', height=20)
        self.tree.heading('date', text='날짜')
        self.tree.heading('shift', text='근무조')
        self.tree.heading('building', text='동')
        self.tree.heading('equipment', text='설비')
        self.tree.heading('phenomenon', text='현상')
        self.tree.heading('cause', text='원인')
        self.tree.heading('action', text='조치')
        self.tree.heading('raw_text', text='기타')
        self.tree.column('date', width=80, anchor='center')
        self.tree.column('shift', width=50, anchor='center')
        self.tree.column('building', width=40, anchor='center')
        self.tree.column('equipment', width=100)
        self.tree.column('phenomenon', width=200)
        self.tree.column('cause', width=200)
        self.tree.column('action', width=200)
        self.tree.column('raw_text', width=250)
        self.tree.tag_configure('new', background='#e8f5e9')

        scrollbar_y = ttk.Scrollbar(right_frame, orient='vertical', command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(right_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_x.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar_y.pack(side='right', fill='y')

        btn_frame = tk.Frame(self.root, padx=10, pady=10)
        btn_frame.pack(fill='x')
        tk.Button(btn_frame, text='신규 전체 저장', width=14, command=self._save, font=('맑은 고딕', 10)).pack(side='right', padx=5)
        tk.Button(btn_frame, text='신규 전체 스킵', width=14, command=self._skip, font=('맑은 고딕', 10)).pack(side='right', padx=5)

        self._load_history_into_records()
        self._refresh_date_list()
        self._update_info()

        if self._pending_dates:
            first_new = sorted(self._pending_dates)[0]
            self._select_date(first_new)
        else:
            self._show_all()

        self.root.mainloop()

    def _load_history_into_records(self):
        if not self.db_path:
            return
        history = get_recent_history(self.db_path, limit=500)
        for rec in history:
            date = rec['date']
            if date not in self._all_records:
                self._all_records[date] = []
            self._all_records[date].append(rec)

    def _refresh_date_list(self):
        if not self.date_listbox:
            return
        self.date_listbox.delete(0, 'end')
        for date in sorted(self._all_records.keys(), reverse=True):
            count = len(self._all_records[date])
            label = f'{date} ({count})'
            if date in self._pending_dates:
                label = f'* {label}'
            self.date_listbox.insert('end', label)
            if date in self._pending_dates:
                idx = self.date_listbox.size() - 1
                self.date_listbox.itemconfig(idx, fg='#2e7d32', selectbackground='#66bb6a')

    def _on_date_select(self, event):
        sel = self.date_listbox.curselection()
        if not sel:
            return
        text = self.date_listbox.get(sel[0])
        date_str = text.lstrip('* ').split(' (')[0]
        self._display_records_for_date(date_str)

    def _select_date(self, date_str):
        if not self.date_listbox:
            return
        for i in range(self.date_listbox.size()):
            text = self.date_listbox.get(i)
            if date_str in text:
                self.date_listbox.selection_clear(0, 'end')
                self.date_listbox.selection_set(i)
                self.date_listbox.see(i)
                self._display_records_for_date(date_str)
                return

    def _display_records_for_date(self, date_str):
        if not self.tree:
            return
        self.tree.delete(*self.tree.get_children())
        records = self._all_records.get(date_str, [])
        is_new = date_str in self._pending_dates
        for rec in records:
            self._insert_record(rec, is_new)

    def _show_all(self):
        if not self.tree:
            return
        self.tree.delete(*self.tree.get_children())
        if self.date_listbox:
            self.date_listbox.selection_clear(0, 'end')
        for date_str in sorted(self._all_records.keys(), reverse=True):
            is_new = date_str in self._pending_dates
            for rec in self._all_records[date_str]:
                self._insert_record(rec, is_new)

    def _insert_record(self, rec, is_new=False):
        phen = _flat(rec.get('phenomenon', '') or '')
        cause = _flat(rec.get('cause', '') or '')
        action = _flat(rec.get('action', '') or '')
        raw = _flat(rec.get('raw_text', '') or '')
        raw_display = '' if (phen or cause or action) else (raw[:80] + '...' if len(raw) > 80 else raw)
        tag = ('new',) if is_new else ()
        self.tree.insert('', 'end', values=(
            rec.get('date', ''), _flat(rec.get('shift', '')), _flat(rec.get('building', '')),
            _flat(rec.get('equipment', '')), phen, cause, action, raw_display
        ), tags=tag)

    def _update_info(self):
        if not self.info_label:
            return
        total_db = sum(len(r) for d, r in self._all_records.items() if d not in self._pending_dates)
        total_new = sum(len(r) for _, _, r, _ in self._pending)
        self.info_label.config(
            text=f'전체 이력: {total_db}건 | 신규 파싱: {total_new}건 ({len(self._pending)}파일)'
        )

    def _save(self):
        self.on_save_all(self._pending)
        self._pending_dates.clear()
        self._pending.clear()
        self._cleanup()

    def _skip(self):
        for _, date_str, records, _ in self._pending:
            if date_str in self._all_records:
                for r in records:
                    if r in self._all_records[date_str]:
                        self._all_records[date_str].remove(r)
                if not self._all_records[date_str]:
                    del self._all_records[date_str]
        self._pending_dates.clear()
        self._pending.clear()
        self._cleanup()

    def _on_close(self):
        self._pending.clear()
        self._pending_dates.clear()
        self._all_records.clear()
        self._cleanup()

    def _cleanup(self):
        self._alive = False
        if self.root:
            self.root.destroy()
            self.root = None
        self.tree = None
        self.date_listbox = None
        self.info_label = None
        with self._lock:
            ParseResultPopup._instance = None


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
        self.last_parse_result = None
        self.icon = None

    def start(self):
        menu = pystray.Menu(
            pystray.MenuItem('파싱 뷰어 열기', self._show_viewer),
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

    def notify(self, message):
        if self.icon:
            self.icon.notify(message, '설비일보 Word 파서')

    def set_last_result(self, doc_name, date_str, records):
        self.last_parse_result = (doc_name, date_str, records)

    def _show_viewer(self, icon, item):
        popup = ParseResultPopup.get_or_create(
            on_save_all=self.save_all_callback,
            db_path=self.db_path
        )
        if not popup._alive:
            threading.Thread(target=popup._show, daemon=True).start()

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
        tray.set_last_result(doc_name, date_str, records)
        tray.notify(f'설비일보 {date_str} 파싱 완료 ({len(records)}건)')
        popup = ParseResultPopup.get_or_create(on_save_all=save_all, db_path=DB_PATH)
        popup.add_records(doc_name, date_str, records, content_hash)

    def on_duplicate_same(doc_name, date_str):
        tray.notify(f'이미 파싱된 파일입니다 ({date_str})')

    def on_duplicate_changed(doc_name, date_str, records, content_hash):
        tray.set_last_result(doc_name, date_str, records)
        tray.notify(f'내용 변경 감지 ({date_str})')
        popup = ParseResultPopup.get_or_create(on_save_all=save_all, db_path=DB_PATH)
        popup.add_records(doc_name, date_str, records, content_hash)

    def on_date_missing(doc_name):
        return ask_date_input(doc_name)

    watcher = WordWatcher(
        db_path=DB_PATH,
        on_new_parse=on_new_parse,
        on_duplicate_same=on_duplicate_same,
        on_duplicate_changed=on_duplicate_changed,
        on_date_missing=on_date_missing,
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
