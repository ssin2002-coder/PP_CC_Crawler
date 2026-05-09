"""
test_parsers.py — 5종 파서 단위 테스트 (COM 모킹)
"""
import pytest
from unittest.mock import MagicMock

from doc_intelligence.parsers import (
    BaseParser,
    ExcelParser,
    WordParser,
    PowerPointParser,
    PdfParser,
    ImageParser,
)
from doc_intelligence.engine import ParsedDocument


# ──────────────────────────────────────────────
# Mock 헬퍼
# ──────────────────────────────────────────────

def _mock_excel():
    app = MagicMock()
    wb = MagicMock()
    ws = MagicMock()
    ws.Name = "Sheet1"
    ws.UsedRange.Rows.Count = 3
    ws.UsedRange.Columns.Count = 2
    ws.UsedRange.Row = 1
    ws.UsedRange.Column = 1

    data = {
        (1, 1): "견적서",
        (1, 2): "",
        (2, 1): "합계",
        (2, 2): 15000000,
        (3, 1): "날짜",
        (3, 2): "2025.03.15",
    }

    def cells(r, c):
        m = MagicMock()
        m.Value = data.get((r, c), "")
        m.MergeCells = r == 1 and c == 1
        m.NumberFormat = "General"
        m.Formula = data.get((r, c), "")
        return m

    ws.Cells = cells
    wb.Worksheets = [ws]
    wb.FullName = "C:\\견적서.xlsx"
    app.ActiveWorkbook = wb
    return app


def _mock_word():
    app = MagicMock()
    doc = MagicMock()

    p1 = MagicMock()
    p1.Range.Text = "제목: 정비비용 정산서"
    p2 = MagicMock()
    p2.Range.Text = "작성일: 2025-03-15"

    doc.Paragraphs = [p1, p2]
    doc.Tables = []
    doc.FullName = "C:\\정산서.docx"
    app.ActiveDocument = doc
    return app


def _mock_ppt():
    app = MagicMock()
    prs = MagicMock()

    shape = MagicMock()
    shape.HasTextFrame = True
    shape.HasTable = False
    shape.TextFrame.TextRange.Text = "슬라이드 제목"

    slide = MagicMock()
    slide.Shapes = [shape]

    prs.Slides = [slide]
    prs.FullName = "C:\\프레젠테이션.pptx"
    app.ActivePresentation = prs
    return app


# ──────────────────────────────────────────────
# BaseParser
# ──────────────────────────────────────────────

def test_base_parser_raises():
    parser = BaseParser()
    with pytest.raises(NotImplementedError):
        parser.parse_from_com(MagicMock())


# ──────────────────────────────────────────────
# ExcelParser
# ──────────────────────────────────────────────

def test_excel_parser_basic():
    parser = ExcelParser()
    result = parser.parse_from_com(_mock_excel())

    assert isinstance(result, ParsedDocument)
    assert result.file_type == "excel"
    assert "견적서" in result.raw_text
    assert "merge_cells" in result.structure


def test_excel_merge_pattern():
    parser = ExcelParser()
    result = parser.parse_from_com(_mock_excel())

    assert "merge_hash" in result.structure


def test_excel_sheet_count():
    parser = ExcelParser()
    result = parser.parse_from_com(_mock_excel())

    assert result.structure.get("sheet_count") == 1
    assert "sheets" in result.structure


def test_excel_cells_populated():
    parser = ExcelParser()
    result = parser.parse_from_com(_mock_excel())

    assert len(result.cells) > 0
    addresses = [c.address for c in result.cells]
    assert any("Sheet1" in addr for addr in addresses)


# ──────────────────────────────────────────────
# WordParser
# ──────────────────────────────────────────────

def test_word_parser():
    parser = WordParser()
    result = parser.parse_from_com(_mock_word())

    assert isinstance(result, ParsedDocument)
    assert result.file_type == "word"
    assert "정비비용" in result.raw_text


def test_word_parser_no_tables():
    parser = WordParser()
    result = parser.parse_from_com(_mock_word())

    assert result.structure.get("table_count") == 0


# ──────────────────────────────────────────────
# PowerPointParser
# ──────────────────────────────────────────────

def test_ppt_parser():
    parser = PowerPointParser()
    result = parser.parse_from_com(_mock_ppt())

    assert isinstance(result, ParsedDocument)
    assert result.file_type == "ppt"
    assert "슬라이드" in result.raw_text


def test_ppt_slide_count():
    parser = PowerPointParser()
    result = parser.parse_from_com(_mock_ppt())

    assert result.structure.get("slide_count") == 1


# ──────────────────────────────────────────────
# PdfParser
# ──────────────────────────────────────────────

def test_pdf_parser_fallback():
    """Acrobat COM이 예외를 발생시키면 fallback 경로로 처리된다."""
    app = MagicMock()
    app.GetActiveDoc.side_effect = Exception("Acrobat not available")

    parser = PdfParser()
    result = parser.parse_from_com(app)

    assert isinstance(result, ParsedDocument)
    assert result.file_type == "pdf"
    assert result.metadata.get("fallback") is True


def test_pdf_parser_fallback_empty_cells():
    """fallback 시 cells는 빈 리스트여야 한다."""
    app = MagicMock()
    app.GetActiveDoc.side_effect = Exception("no acrobat")

    parser = PdfParser()
    result = parser.parse_from_com(app)

    assert result.cells == []


# ──────────────────────────────────────────────
# ImageParser
# ──────────────────────────────────────────────

def test_image_parser_no_tesseract():
    """존재하지 않는 파일 경로 전달 시 graceful fail."""
    parser = ImageParser()
    result = parser.parse_from_com("C:\\nonexistent_image_12345.png")

    assert isinstance(result, ParsedDocument)
    assert result.file_type == "image"


def test_image_parser_returns_parsed_document():
    """ImageParser는 항상 ParsedDocument를 반환해야 한다."""
    parser = ImageParser()
    result = parser.parse_from_com("C:\\fake_path.jpg")

    assert hasattr(result, "file_path")
    assert hasattr(result, "raw_text")
    assert hasattr(result, "cells")
