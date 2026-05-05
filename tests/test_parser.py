# tests/test_parser.py
import pytest
from word_crawler import (
    clean_cell_text, split_items, extract_date_from_text,
    extract_date_from_filename, find_main_table_index, parse_table_data,
    parse_item_block,
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

    def test_from_text_ilja_label(self):
        assert extract_date_from_text('일자 : 2024년 5월 3일') == '2024-05-03'

    def test_from_text_ilja_label_fullwidth_colon(self):
        assert extract_date_from_text('일자 ：2024년 12월 31일') == '2024-12-31'

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


class TestParseItemBlock:
    def test_title_and_dash_body(self):
        text = '*AHU-3 이상진동\n - 베어링 발주\n - 임시조치'
        parsed = parse_item_block(text)
        assert parsed['title'] == 'AHU-3 이상진동'
        assert parsed['raw_text'] == '- 베어링 발주\n- 임시조치'

    def test_title_only(self):
        parsed = parse_item_block('*제목만 있음')
        assert parsed['title'] == '제목만 있음'
        assert parsed['raw_text'] == ''

    def test_no_title_marker(self):
        parsed = parse_item_block('단순 본문\n계속')
        assert parsed['title'] == ''
        assert parsed['raw_text'] == '단순 본문\n계속'

    def test_strips_numbering(self):
        parsed = parse_item_block('1) *AHU 이상\n2) - 점검')
        assert parsed['title'] == 'AHU 이상'
        assert parsed['raw_text'] == '- 점검'


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
        assert a_rec[0]['raw_text'] == 'AHU-3 이상진동'
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
        assert a_recs[0]['raw_text'] == 'AHU 이상'
        assert a_recs[1]['raw_text'] == '보일러 점검'
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
