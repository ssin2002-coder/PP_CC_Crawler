"""
정비 비용정산 Excel 크롤러 (단일 파일)
- 열린 Excel 문서에서 정비 비용정산 시트를 자동 감지/파싱
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

DEPARTMENTS = ["인프라팀", "설비팀", "유틸팀", "전기팀", "안전팀"]
MAINTENANCE_TYPES = ["예방정비", "사후정비", "개선정비", "긴급정비"]
EXPENSE_TYPES = ["운반비", "장비임대", "폐기물처리", "기타"]
WORKER_TYPES = ["내부", "외주"]

EQUIPMENT_MASTER = {
    "CH-001": {"name": "Chiller #1", "type": "Chiller", "location": "유틸동 B1"},
    "CH-002": {"name": "Chiller #2", "type": "Chiller", "location": "유틸동 B1"},
    "CH-003": {"name": "Chiller #3", "type": "Chiller", "location": "유틸동 B1"},
    "AHU-001": {"name": "AHU #1", "type": "AHU", "location": "클린룸 2F"},
    "AHU-002": {"name": "AHU #2", "type": "AHU", "location": "클린룸 3F"},
    "AHU-003": {"name": "AHU #3", "type": "AHU", "location": "클린룸 4F"},
    "PUMP-001": {"name": "PCW Pump #1", "type": "Pump", "location": "유틸동 1F"},
    "PUMP-002": {"name": "PCW Pump #2", "type": "Pump", "location": "유틸동 1F"},
    "CT-001": {"name": "Cooling Tower #1", "type": "CoolingTower", "location": "옥상"},
    "CT-002": {"name": "Cooling Tower #2", "type": "CoolingTower", "location": "옥상"},
    "BLR-001": {"name": "Boiler #1", "type": "Boiler", "location": "유틸동 1F"},
    "CDA-001": {"name": "CDA Compressor #1", "type": "Compressor", "location": "유틸동 B1"},
}

# 정산 시트 식별용 헤더 키워드
SETTLEMENT_HEADER_KEYWORDS = ['정산', '비용', '정비']
MATERIAL_HEADER_KEYWORDS = ['자재명', '규격', '수량', '단가', '금액']
LABOR_HEADER_KEYWORDS = ['구분', '성명', '시간', '단가', '금액']
EXPENSE_HEADER_KEYWORDS = ['구분', '내용', '금액']

DATE_PATTERNS = [
    re.compile(r'(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})'),
    re.compile(r'(\d{4})[\-.](\d{2})[\-.](\d{2})'),
]

DATE_PATTERNS_FILENAME = [
    re.compile(r'(\d{4})(\d{2})(\d{2})'),
    re.compile(r'(\d{4})[\-_.](\d{2})[\-_.](\d{2})'),
]

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'settlement.db')
LOCK_PATH = os.path.join(os.path.dirname(DB_PATH), 'excel_crawler.lock')

# ─────────────────────────────────────────────
# DB
# ─────────────────────────────────────────────

def init_db(db_path=None):
    path = db_path or DB_PATH
    os.makedirs(os.path.dirname(path), exist_ok=True)
    conn = sqlite3.connect(path)
    conn.execute('''
        CREATE TABLE IF NOT EXISTS settlement_header (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            settlement_date TEXT NOT NULL,
            settlement_month TEXT NOT NULL,
            author          TEXT,
            department      TEXT,
            maintenance_type TEXT,
            equipment_code  TEXT,
            equipment_name  TEXT,
            material_total  REAL DEFAULT 0,
            labor_total     REAL DEFAULT 0,
            expense_total   REAL DEFAULT 0,
            grand_total     REAL DEFAULT 0,
            remarks         TEXT,
            source_file     TEXT NOT NULL,
            content_hash    TEXT,
            created_at      TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS settlement_material (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            header_id     INTEGER NOT NULL REFERENCES settlement_header(id),
            part_name     TEXT NOT NULL,
            specification TEXT,
            quantity      REAL DEFAULT 0,
            unit_price    REAL DEFAULT 0,
            amount        REAL DEFAULT 0
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS settlement_labor (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            header_id     INTEGER NOT NULL REFERENCES settlement_header(id),
            worker_type   TEXT NOT NULL,
            worker_name   TEXT,
            hours         REAL DEFAULT 0,
            hourly_rate   REAL DEFAULT 0,
            amount        REAL DEFAULT 0
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS settlement_expense (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            header_id     INTEGER NOT NULL REFERENCES settlement_header(id),
            expense_type  TEXT NOT NULL,
            description   TEXT,
            amount        REAL DEFAULT 0
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_sh_date ON settlement_header(settlement_date)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_sh_equip ON settlement_header(equipment_code)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_sh_source ON settlement_header(source_file)')
    conn.commit()
    conn.close()


def insert_settlement(db_path, header, materials, labors, expenses, content_hash=None):
    conn = sqlite3.connect(db_path)
    cur = conn.execute('''
        INSERT INTO settlement_header
        (settlement_date, settlement_month, author, department, maintenance_type,
         equipment_code, equipment_name, material_total, labor_total, expense_total,
         grand_total, remarks, source_file, content_hash)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        header['settlement_date'], header['settlement_month'],
        header.get('author', ''), header.get('department', ''),
        header.get('maintenance_type', ''), header.get('equipment_code', ''),
        header.get('equipment_name', ''),
        header.get('material_total', 0), header.get('labor_total', 0),
        header.get('expense_total', 0), header.get('grand_total', 0),
        header.get('remarks', ''), header['source_file'], content_hash
    ))
    header_id = cur.lastrowid

    for m in materials:
        if m.get('part_name'):
            conn.execute('''
                INSERT INTO settlement_material (header_id, part_name, specification, quantity, unit_price, amount)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (header_id, m['part_name'], m.get('specification', ''),
                  m.get('quantity', 0), m.get('unit_price', 0), m.get('amount', 0)))

    for l in labors:
        if l.get('worker_type'):
            conn.execute('''
                INSERT INTO settlement_labor (header_id, worker_type, worker_name, hours, hourly_rate, amount)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (header_id, l['worker_type'], l.get('worker_name', ''),
                  l.get('hours', 0), l.get('hourly_rate', 0), l.get('amount', 0)))

    for e in expenses:
        if e.get('expense_type'):
            conn.execute('''
                INSERT INTO settlement_expense (header_id, expense_type, description, amount)
                VALUES (?, ?, ?, ?)
            ''', (header_id, e['expense_type'], e.get('description', ''), e.get('amount', 0)))

    conn.commit()
    conn.close()
    return header_id


def check_duplicate(db_path, source_file, settlement_date, new_hash):
    conn = sqlite3.connect(db_path)
    row = conn.execute(
        'SELECT content_hash FROM settlement_header WHERE source_file = ? AND settlement_date = ? LIMIT 1',
        (source_file, settlement_date)
    ).fetchone()
    conn.close()
    if row is None:
        return 'new'
    return 'same' if row[0] == new_hash else 'changed'


def delete_by_source(db_path, source_file, settlement_date):
    conn = sqlite3.connect(db_path)
    rows = conn.execute(
        'SELECT id FROM settlement_header WHERE source_file = ? AND settlement_date = ?',
        (source_file, settlement_date)
    ).fetchall()
    for (hid,) in rows:
        conn.execute('DELETE FROM settlement_material WHERE header_id = ?', (hid,))
        conn.execute('DELETE FROM settlement_labor WHERE header_id = ?', (hid,))
        conn.execute('DELETE FROM settlement_expense WHERE header_id = ?', (hid,))
    conn.execute(
        'DELETE FROM settlement_header WHERE source_file = ? AND settlement_date = ?',
        (source_file, settlement_date)
    )
    conn.commit()
    conn.close()


def get_recent_history(db_path, limit=500):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        'SELECT * FROM settlement_header ORDER BY settlement_date DESC, id DESC LIMIT ?',
        (limit,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def compute_hash(header, materials, labors, expenses):
    parts = [
        header.get('settlement_date', ''),
        header.get('equipment_code', ''),
        str(header.get('grand_total', 0)),
    ]
    for m in materials:
        parts.append(f"M:{m.get('part_name', '')}:{m.get('amount', 0)}")
    for l in labors:
        parts.append(f"L:{l.get('worker_type', '')}:{l.get('amount', 0)}")
    for e in expenses:
        parts.append(f"E:{e.get('expense_type', '')}:{e.get('amount', 0)}")
    raw = '|'.join(parts)
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()[:16]

# ─────────────────────────────────────────────
# Parser
# ─────────────────────────────────────────────

def _safe_str(val):
    if val is None:
        return ''
    return str(val).strip()


def _safe_float(val):
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        cleaned = re.sub(r'[^\d.\-]', '', str(val))
        try:
            return float(cleaned) if cleaned else 0.0
        except ValueError:
            return 0.0


def extract_date(text):
    for pattern in DATE_PATTERNS:
        match = pattern.search(str(text))
        if match:
            y, m, d = int(match.group(1)), int(match.group(2)), int(match.group(3))
            if y < 100:
                y += 2000
            return f'{y:04d}-{m:02d}-{d:02d}'
    return None


def extract_date_from_filename(filename):
    for pattern in DATE_PATTERNS_FILENAME:
        match = pattern.search(filename)
        if match:
            y, m, d = int(match.group(1)), int(match.group(2)), int(match.group(3))
            if y < 100:
                y += 2000
            return f'{y:04d}-{m:02d}-{d:02d}'
    return None


def _find_cell_value(sheet, search_text, max_row=30, max_col=10):
    """시트에서 특정 텍스트를 포함한 셀의 오른쪽 값을 반환"""
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            val = sheet.Cells(r, c).Value
            if val and search_text in str(val):
                right_val = sheet.Cells(r, c + 1).Value
                return _safe_str(right_val), r, c
    return None, None, None


def _find_table_start(sheet, keywords, max_row=50, max_col=15):
    """헤더 키워드가 포함된 행을 찾아 테이블 시작 위치 반환"""
    for r in range(1, max_row + 1):
        row_texts = []
        for c in range(1, max_col + 1):
            val = sheet.Cells(r, c).Value
            row_texts.append(_safe_str(val).replace(' ', ''))
        combined = ' '.join(row_texts)
        match_count = sum(1 for kw in keywords if kw in combined)
        if match_count >= len(keywords) - 1 and match_count > 0:
            return r
    return None


def _detect_equipment_code(text):
    """텍스트에서 설비 코드를 매칭"""
    text_upper = str(text).upper().replace(' ', '')
    for code in EQUIPMENT_MASTER:
        if code.upper().replace('-', '') in text_upper.replace('-', ''):
            return code
    return ''


def parse_excel_sheet(sheet, workbook_name):
    """Excel 시트에서 정산 데이터 파싱 → (header, materials, labors, expenses) or None"""

    # 1) 정산 시트인지 확인
    sheet_name = sheet.Name
    all_text = ''
    for r in range(1, 15):
        for c in range(1, 10):
            val = sheet.Cells(r, c).Value
            if val:
                all_text += str(val) + ' '

    is_settlement = any(kw in all_text for kw in SETTLEMENT_HEADER_KEYWORDS)
    if not is_settlement:
        return None

    # 2) 헤더 정보 추출
    date_str = None
    author = ''
    department = ''
    maint_type = ''
    equip_text = ''
    remarks = ''

    # 날짜 찾기
    for label in ['정산일', '날짜', '일자', '작업일']:
        val, _, _ = _find_cell_value(sheet, label)
        if val:
            date_str = extract_date(val)
            if date_str:
                break

    if not date_str:
        date_str = extract_date(all_text)
    if not date_str:
        date_str = extract_date_from_filename(workbook_name)

    # 작성자/부서/정비유형/설비
    for label in ['작성자', '담당자', '작성']:
        val, _, _ = _find_cell_value(sheet, label)
        if val:
            author = val
            break

    for label in ['부서', '팀']:
        val, _, _ = _find_cell_value(sheet, label)
        if val:
            department = val
            break

    for label in ['정비유형', '정비구분', '유형']:
        val, _, _ = _find_cell_value(sheet, label)
        if val:
            maint_type = val
            break

    for label in ['설비', '장비', '대상설비', '설비코드']:
        val, _, _ = _find_cell_value(sheet, label)
        if val:
            equip_text = val
            break

    for label in ['비고', '메모', '특이사항']:
        val, _, _ = _find_cell_value(sheet, label)
        if val:
            remarks = val
            break

    equip_code = _detect_equipment_code(equip_text)
    equip_name = EQUIPMENT_MASTER[equip_code]['name'] if equip_code in EQUIPMENT_MASTER else equip_text

    # 3) 자재비 테이블 파싱
    materials = []
    mat_row = _find_table_start(sheet, ['자재', '수량', '금액'])
    if mat_row:
        for r in range(mat_row + 1, mat_row + 50):
            part_name = _safe_str(sheet.Cells(r, 1).Value)
            if not part_name:
                # 다음 행도 비어있으면 테이블 끝
                if not _safe_str(sheet.Cells(r + 1, 1).Value):
                    break
                continue
            # 소계/합계 행이면 중단
            if any(kw in part_name for kw in ['소계', '합계', '계']):
                break
            materials.append({
                'part_name': part_name,
                'specification': _safe_str(sheet.Cells(r, 2).Value),
                'quantity': _safe_float(sheet.Cells(r, 3).Value),
                'unit_price': _safe_float(sheet.Cells(r, 4).Value),
                'amount': _safe_float(sheet.Cells(r, 5).Value),
            })

    # 4) 인건비 테이블 파싱
    labors = []
    labor_row = _find_table_start(sheet, ['인건', '시간', '금액'])
    if not labor_row:
        labor_row = _find_table_start(sheet, ['작업자', '시간'])
    if labor_row:
        for r in range(labor_row + 1, labor_row + 50):
            worker_type = _safe_str(sheet.Cells(r, 1).Value)
            if not worker_type:
                if not _safe_str(sheet.Cells(r + 1, 1).Value):
                    break
                continue
            if any(kw in worker_type for kw in ['소계', '합계', '계']):
                break
            labors.append({
                'worker_type': worker_type,
                'worker_name': _safe_str(sheet.Cells(r, 2).Value),
                'hours': _safe_float(sheet.Cells(r, 3).Value),
                'hourly_rate': _safe_float(sheet.Cells(r, 4).Value),
                'amount': _safe_float(sheet.Cells(r, 5).Value),
            })

    # 5) 경비 테이블 파싱
    expenses = []
    exp_row = _find_table_start(sheet, ['경비', '금액'])
    if not exp_row:
        exp_row = _find_table_start(sheet, ['기타', '내용', '금액'])
    if exp_row:
        for r in range(exp_row + 1, exp_row + 50):
            exp_type = _safe_str(sheet.Cells(r, 1).Value)
            if not exp_type:
                if not _safe_str(sheet.Cells(r + 1, 1).Value):
                    break
                continue
            if any(kw in exp_type for kw in ['소계', '합계', '계']):
                break
            expenses.append({
                'expense_type': exp_type,
                'description': _safe_str(sheet.Cells(r, 2).Value),
                'amount': _safe_float(sheet.Cells(r, 3).Value),
            })

    # 데이터가 하나도 없으면 스킵
    if not materials and not labors and not expenses:
        return None

    material_total = sum(m.get('amount', 0) for m in materials)
    labor_total = sum(l.get('amount', 0) for l in labors)
    expense_total = sum(e.get('amount', 0) for e in expenses)

    header = {
        'settlement_date': date_str or '',
        'settlement_month': date_str[:7] if date_str and len(date_str) >= 7 else '',
        'author': author,
        'department': department,
        'maintenance_type': maint_type,
        'equipment_code': equip_code,
        'equipment_name': equip_name,
        'material_total': material_total,
        'labor_total': labor_total,
        'expense_total': expense_total,
        'grand_total': material_total + labor_total + expense_total,
        'remarks': remarks,
        'source_file': workbook_name,
    }

    return header, materials, labors, expenses

# ─────────────────────────────────────────────
# Excel Watcher
# ─────────────────────────────────────────────

class ExcelWatcher:
    def __init__(self, db_path, on_new_parse, on_duplicate_same, on_duplicate_changed, on_date_missing):
        self.db_path = db_path
        self.on_new_parse = on_new_parse
        self.on_duplicate_same = on_duplicate_same
        self.on_duplicate_changed = on_duplicate_changed
        self.on_date_missing = on_date_missing
        self._stop_event = threading.Event()
        self._seen_books = set()
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
                    self._check_excel()
                except Exception:
                    pass
                self._stop_event.wait(4)
        finally:
            pythoncom.CoUninitialize()

    def _check_excel(self):
        try:
            excel = win32com.client.GetActiveObject('Excel.Application')
        except Exception:
            return

        for i in range(1, excel.Workbooks.Count + 1):
            wb = excel.Workbooks(i)
            wb_name = wb.Name
            wb_full = wb.FullName
            if wb_full in self._seen_books:
                continue

            results = self._parse_workbook(wb)
            if not results:
                continue

            self._seen_books.add(wb_full)

            for header, materials, labors, expenses in results:
                date_str = header['settlement_date']
                if not date_str:
                    date_str = self.on_date_missing(wb_name)
                    if not date_str:
                        continue
                    header['settlement_date'] = date_str
                    header['settlement_month'] = date_str[:7] if len(date_str) >= 7 else ''

                content_hash = compute_hash(header, materials, labors, expenses)
                dup_status = check_duplicate(self.db_path, wb_name, date_str, content_hash)

                if dup_status == 'new':
                    self.on_new_parse(wb_name, date_str, header, materials, labors, expenses, content_hash)
                elif dup_status == 'same':
                    self.on_duplicate_same(wb_name, date_str)
                else:
                    self.on_duplicate_changed(wb_name, date_str, header, materials, labors, expenses, content_hash)

    def _parse_workbook(self, wb):
        results = []
        for i in range(1, wb.Sheets.Count + 1):
            sheet = wb.Sheets(i)
            try:
                parsed = parse_excel_sheet(sheet, wb.Name)
                if parsed:
                    results.append(parsed)
            except Exception:
                pass
        return results

    def reset_seen(self):
        self._seen_books.clear()

# ─────────────────────────────────────────────
# UI (tkinter)
# ─────────────────────────────────────────────

def _flat(text):
    if not text:
        return ''
    return ' '.join(str(text).replace('\r', ' ').replace('\n', ' ').split())


def _fmt_money(val):
    try:
        return f'{int(val):,}' if val else '0'
    except (ValueError, TypeError):
        return str(val)


class SettlementPopup:
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
        self._pending = []  # [(wb_name, date, header, materials, labors, expenses, hash), ...]
        self._all_records = {}  # {date: [header_dicts...]}
        self._pending_dates = set()
        self._alive = False
        self.root = None
        self.tree = None
        self.date_listbox = None
        self.info_label = None

    def add_records(self, wb_name, date_str, header, materials, labors, expenses, content_hash):
        self._pending.append((wb_name, date_str, header, materials, labors, expenses, content_hash))
        self._pending_dates.add(date_str)
        if date_str not in self._all_records:
            self._all_records[date_str] = []
        self._all_records[date_str].append(header)

        if self._alive and self.root:
            self.root.after(0, self._refresh_date_list)
            self.root.after(0, lambda: self._select_date(date_str))
            self.root.after(0, self._update_info)
        else:
            threading.Thread(target=self._show, daemon=True).start()

    def _show(self):
        self._alive = True
        self.root = tk.Tk()
        self.root.title('정비 비용정산 파서 (Excel)')
        self.root.geometry('1200x550')
        self.root.attributes('-topmost', True)
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

        info_frame = tk.Frame(self.root, padx=10, pady=5)
        info_frame.pack(fill='x')
        self.info_label = tk.Label(info_frame, text='', anchor='w', font=('맑은 고딕', 10))
        self.info_label.pack(fill='x')

        main_frame = tk.Frame(self.root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # Left: date list
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

        # Right: table
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side='left', fill='both', expand=True)

        columns = ('date', 'department', 'maint_type', 'equip', 'material', 'labor', 'expense', 'total', 'source')
        self.tree = ttk.Treeview(right_frame, columns=columns, show='headings', height=18)
        self.tree.heading('date', text='정산일')
        self.tree.heading('department', text='부서')
        self.tree.heading('maint_type', text='정비유형')
        self.tree.heading('equip', text='설비')
        self.tree.heading('material', text='자재비')
        self.tree.heading('labor', text='인건비')
        self.tree.heading('expense', text='경비')
        self.tree.heading('total', text='합계')
        self.tree.heading('source', text='파일')
        self.tree.column('date', width=90, anchor='center')
        self.tree.column('department', width=70, anchor='center')
        self.tree.column('maint_type', width=70, anchor='center')
        self.tree.column('equip', width=120)
        self.tree.column('material', width=100, anchor='e')
        self.tree.column('labor', width=100, anchor='e')
        self.tree.column('expense', width=100, anchor='e')
        self.tree.column('total', width=110, anchor='e')
        self.tree.column('source', width=150)
        self.tree.tag_configure('new', background='#e8f5e9')

        scrollbar_y = ttk.Scrollbar(right_frame, orient='vertical', command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(right_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_x.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar_y.pack(side='right', fill='y')

        # Buttons
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
            date = rec['settlement_date']
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
        tag = ('new',) if is_new else ()
        equip_display = rec.get('equipment_name', '') or rec.get('equipment_code', '')
        self.tree.insert('', 'end', values=(
            rec.get('settlement_date', ''),
            _flat(rec.get('department', '')),
            _flat(rec.get('maintenance_type', '')),
            _flat(equip_display),
            _fmt_money(rec.get('material_total', 0)),
            _fmt_money(rec.get('labor_total', 0)),
            _fmt_money(rec.get('expense_total', 0)),
            _fmt_money(rec.get('grand_total', 0)),
            _flat(rec.get('source_file', '')),
        ), tags=tag)

    def _update_info(self):
        if not self.info_label:
            return
        total_db = sum(len(r) for d, r in self._all_records.items() if d not in self._pending_dates)
        total_new = len(self._pending)
        self.info_label.config(
            text=f'전체 이력: {total_db}건 | 신규 파싱: {total_new}건'
        )

    def _save(self):
        self.on_save_all(self._pending)
        self._pending_dates.clear()
        self._pending.clear()
        self._cleanup()

    def _skip(self):
        for _, date_str, header, *_ in self._pending:
            if date_str in self._all_records:
                if header in self._all_records[date_str]:
                    self._all_records[date_str].remove(header)
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
            SettlementPopup._instance = None


def ask_date_input(wb_name):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    date_str = simpledialog.askstring(
        '날짜 입력',
        f'"{wb_name}"에서 정산일을 찾을 수 없습니다.\n날짜를 입력해 주십시오 (예: 2024-05-03):',
        parent=root
    )
    root.destroy()
    return date_str

# ─────────────────────────────────────────────
# Tray
# ─────────────────────────────────────────────

def create_icon_image():
    img = Image.new('RGB', (64, 64), color=(34, 139, 34))
    draw = ImageDraw.Draw(img)
    # 스프레드시트 아이콘
    draw.rectangle([14, 10, 50, 54], outline='white', width=2)
    draw.line([14, 22, 50, 22], fill='white', width=1)
    draw.line([14, 34, 50, 34], fill='white', width=1)
    draw.line([14, 46, 50, 46], fill='white', width=1)
    draw.line([32, 10, 32, 54], fill='white', width=1)
    return img


class TrayApp:
    def __init__(self, db_path, save_all_callback=None):
        self.db_path = db_path
        self.save_all_callback = save_all_callback or (lambda p: None)
        self.icon = None

    def start(self):
        menu = pystray.Menu(
            pystray.MenuItem('정산 뷰어 열기', self._show_viewer),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('종료', self._quit),
        )
        self.icon = pystray.Icon(
            '정비 비용정산 Excel 파서',
            create_icon_image(),
            '정비 비용정산 Excel 파서',
            menu
        )
        self.icon.run()

    def notify(self, message):
        if self.icon:
            self.icon.notify(message, '정비 비용정산 Excel 파서')

    def _show_viewer(self, icon, item):
        popup = SettlementPopup.get_or_create(
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
        for wb_name, date_str, header, materials, labors, expenses, content_hash in pending:
            delete_by_source(DB_PATH, wb_name, date_str)
            insert_settlement(DB_PATH, header, materials, labors, expenses, content_hash)
        tray.notify(f'전체 저장 완료 ({len(pending)}건)')

    tray = TrayApp(DB_PATH, save_all_callback=save_all)

    def on_new_parse(wb_name, date_str, header, materials, labors, expenses, content_hash):
        total = header.get('grand_total', 0)
        tray.notify(f'정산 {date_str} 파싱 완료 (합계: {_fmt_money(total)}원)')
        popup = SettlementPopup.get_or_create(on_save_all=save_all, db_path=DB_PATH)
        popup.add_records(wb_name, date_str, header, materials, labors, expenses, content_hash)

    def on_duplicate_same(wb_name, date_str):
        tray.notify(f'이미 파싱된 파일입니다 ({date_str})')

    def on_duplicate_changed(wb_name, date_str, header, materials, labors, expenses, content_hash):
        tray.notify(f'내용 변경 감지 ({date_str})')
        popup = SettlementPopup.get_or_create(on_save_all=save_all, db_path=DB_PATH)
        popup.add_records(wb_name, date_str, header, materials, labors, expenses, content_hash)

    def on_date_missing(wb_name):
        return ask_date_input(wb_name)

    watcher = ExcelWatcher(
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
