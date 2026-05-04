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
