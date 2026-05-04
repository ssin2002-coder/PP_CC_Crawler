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
