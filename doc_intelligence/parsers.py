"""
parsers.py — 5종 문서 파서 (Excel / Word / PPT / PDF / Image)

모든 파서는 COM을 통해 사용자가 열어놓은 파일에 접근한다.
"""
import hashlib
import logging

from doc_intelligence.engine import ParsedDocument, CellData

logger = logging.getLogger(__name__)

# pytesseract는 미설치 환경에서도 import 성공해야 한다.
try:
    import pytesseract
    from PIL import Image as PILImage
    _TESSERACT_AVAILABLE = True
except ImportError:
    _TESSERACT_AVAILABLE = False


# ──────────────────────────────────────────────
# BaseParser
# ──────────────────────────────────────────────

class BaseParser:
    """모든 파서의 베이스 클래스."""

    def parse_from_com(self, com_app) -> ParsedDocument:
        raise NotImplementedError


# ──────────────────────────────────────────────
# ExcelParser
# ──────────────────────────────────────────────

class ExcelParser(BaseParser):
    """Excel.Application COM에서 ActiveWorkbook을 파싱한다."""

    def parse_from_com(self, com_app) -> ParsedDocument:
        wb = com_app.ActiveWorkbook
        file_path = str(wb.FullName)

        all_cells: list[CellData] = []
        raw_parts: list[str] = []
        sheets_meta: list[dict] = []
        merge_cells: list[str] = []

        for ws in wb.Worksheets:
            sheet_name = ws.Name
            row_start = ws.UsedRange.Row
            col_start = ws.UsedRange.Column
            row_count = ws.UsedRange.Rows.Count
            col_count = ws.UsedRange.Columns.Count

            row_end = row_start + row_count - 1
            col_end = col_start + col_count - 1

            for r in range(row_start, row_end + 1):
                for c in range(col_start, col_end + 1):
                    cell = ws.Cells(r, c)
                    value = cell.Value
                    is_merged = bool(cell.MergeCells)
                    fmt = str(cell.NumberFormat) if cell.NumberFormat else "General"
                    formula = cell.Formula

                    # data_type 판별
                    if formula and str(formula).startswith("="):
                        data_type = "formula"
                    elif isinstance(value, (int, float)):
                        data_type = "number"
                    elif value and any(ch in str(value) for ch in (".", "-", "/")) and _looks_like_date(str(value)):
                        data_type = "date"
                    else:
                        data_type = "text"

                    address = f"{sheet_name}!R{r}C{c}"

                    if is_merged:
                        merge_cells.append(address)

                    neighbors = {
                        "merged": is_merged,
                        "number_format": fmt,
                        "row": r,
                        "col": c,
                    }

                    cell_data = CellData(
                        address=address,
                        value=value,
                        data_type=data_type,
                        neighbors=neighbors,
                    )
                    all_cells.append(cell_data)

                    if value is not None and str(value).strip():
                        raw_parts.append(str(value))

            sheets_meta.append({
                "name": sheet_name,
                "rows": row_count,
                "cols": col_count,
            })

        # merge_hash: 병합 셀 패턴의 MD5
        merge_pattern_str = "|".join(sorted(merge_cells))
        merge_hash = hashlib.md5(merge_pattern_str.encode("utf-8")).hexdigest()

        structure = {
            "sheets": sheets_meta,
            "sheet_count": len(sheets_meta),
            "merge_cells": merge_cells,
            "merge_hash": merge_hash,
        }

        return ParsedDocument(
            file_path=file_path,
            file_type="excel",
            raw_text="\n".join(raw_parts),
            structure=structure,
            cells=all_cells,
            metadata={},
        )


def _looks_like_date(value: str) -> bool:
    """단순 휴리스틱: 숫자와 구분자(. - /)로만 구성된 문자열이면 날짜로 간주."""
    import re
    return bool(re.match(r"^\d[\d.\-/]+\d$", value.strip()))


# ──────────────────────────────────────────────
# WordParser
# ──────────────────────────────────────────────

class WordParser(BaseParser):
    """Word.Application COM에서 ActiveDocument를 파싱한다."""

    def parse_from_com(self, com_app) -> ParsedDocument:
        doc = com_app.ActiveDocument
        file_path = str(doc.FullName)

        all_cells: list[CellData] = []
        raw_parts: list[str] = []

        # Paragraphs 순회
        for i, para in enumerate(doc.Paragraphs):
            text = str(para.Range.Text).rstrip("\r\n")
            if text.strip():
                raw_parts.append(text)
                all_cells.append(CellData(
                    address=f"para:{i}",
                    value=text,
                    data_type="text",
                    neighbors={"index": i},
                ))

        # Tables 순회
        table_count = 0
        for t_idx, table in enumerate(doc.Tables):
            table_count += 1
            for r_idx, row in enumerate(table.Rows):
                for c_idx, cell in enumerate(row.Cells):
                    text = str(cell.Range.Text).rstrip("\r\n\x07")
                    if text.strip():
                        raw_parts.append(text)
                    address = f"table{t_idx}:R{r_idx}C{c_idx}"
                    all_cells.append(CellData(
                        address=address,
                        value=text,
                        data_type="text",
                        neighbors={"table": t_idx, "row": r_idx, "col": c_idx},
                    ))

        structure = {
            "paragraph_count": len(doc.Paragraphs),
            "table_count": table_count,
        }

        return ParsedDocument(
            file_path=file_path,
            file_type="word",
            raw_text="\n".join(raw_parts),
            structure=structure,
            cells=all_cells,
            metadata={},
        )


# ──────────────────────────────────────────────
# PowerPointParser
# ──────────────────────────────────────────────

class PowerPointParser(BaseParser):
    """PowerPoint.Application COM에서 ActivePresentation을 파싱한다."""

    def parse_from_com(self, com_app) -> ParsedDocument:
        prs = com_app.ActivePresentation
        file_path = str(prs.FullName)

        all_cells: list[CellData] = []
        raw_parts: list[str] = []

        for s_idx, slide in enumerate(prs.Slides):
            slide_num = s_idx + 1
            for sh_idx, shape in enumerate(slide.Shapes):
                shape_num = sh_idx + 1

                # TextFrame 처리
                if shape.HasTextFrame:
                    text = str(shape.TextFrame.TextRange.Text).strip()
                    if text:
                        raw_parts.append(text)
                    address = f"slide{slide_num}:shape{shape_num}"
                    all_cells.append(CellData(
                        address=address,
                        value=text,
                        data_type="text",
                        neighbors={"slide": slide_num, "shape": shape_num},
                    ))

                # Table 처리
                if shape.HasTable:
                    for r_idx, row in enumerate(shape.Table.Rows):
                        for c_idx, cell in enumerate(row.Cells):
                            try:
                                text = str(cell.Shape.TextFrame.TextRange.Text).strip()
                            except Exception:
                                text = ""
                            if text:
                                raw_parts.append(text)
                            address = f"slide{slide_num}:tbl:R{r_idx}C{c_idx}"
                            all_cells.append(CellData(
                                address=address,
                                value=text,
                                data_type="text",
                                neighbors={"slide": slide_num, "row": r_idx, "col": c_idx},
                            ))

        structure = {
            "slide_count": len(prs.Slides),
        }

        return ParsedDocument(
            file_path=file_path,
            file_type="ppt",
            raw_text="\n".join(raw_parts),
            structure=structure,
            cells=all_cells,
            metadata={},
        )


# ──────────────────────────────────────────────
# PdfParser
# ──────────────────────────────────────────────

class PdfParser(BaseParser):
    """AcroExch.App COM으로 PDF 텍스트를 추출한다.

    Acrobat Pro 미설치(또는 COM 예외) 시 _fallback_ocr를 호출한다.
    """

    def parse_from_com(self, com_app) -> ParsedDocument:
        try:
            return self._parse_acrobat(com_app)
        except Exception as exc:
            logger.warning("Acrobat COM 파싱 실패, fallback 전환: %s", exc)
            return self._fallback_ocr(com_app)

    def _parse_acrobat(self, com_app) -> ParsedDocument:
        doc = com_app.GetActiveDoc()
        js = doc.GetJSObject()
        page_count = doc.GetNumPages()

        all_cells: list[CellData] = []
        raw_parts: list[str] = []

        for page_idx in range(page_count):
            page = doc.GetNthPage(page_idx)
            word_count = page.GetNumWords()
            for word_idx in range(word_count):
                word = js.getPageNthWord(page_idx, word_idx)
                if word and str(word).strip():
                    text = str(word).strip()
                    raw_parts.append(text)
                    address = f"pdf:p{page_idx}w{word_idx}"
                    all_cells.append(CellData(
                        address=address,
                        value=text,
                        data_type="text",
                        neighbors={"page": page_idx, "word_index": word_idx},
                    ))

        file_path = ""
        try:
            file_path = str(doc.GetFileName())
        except Exception:
            pass

        return ParsedDocument(
            file_path=file_path,
            file_type="pdf",
            raw_text="\n".join(raw_parts),
            structure={"page_count": page_count},
            cells=all_cells,
            metadata={},
        )

    def _fallback_ocr(self, com_app) -> ParsedDocument:
        """Acrobat 불가 시 빈 문서를 반환하고 fallback=True 플래그를 설정한다."""
        return ParsedDocument(
            file_path="",
            file_type="pdf",
            raw_text="",
            structure={},
            cells=[],
            metadata={"fallback": True},
        )


# ──────────────────────────────────────────────
# ImageParser
# ──────────────────────────────────────────────

class ImageParser(BaseParser):
    """pytesseract OCR로 이미지 파일을 파싱한다.

    com_app 인자로 이미지 파일 경로(str)를 받는다.
    pytesseract 미설치 또는 파일 접근 실패 시 graceful fail.
    """

    def parse_from_com(self, com_app) -> ParsedDocument:
        file_path = str(com_app) if com_app else ""

        if not _TESSERACT_AVAILABLE:
            logger.warning("pytesseract 미설치 — ImageParser graceful fail")
            return _empty_image_doc(file_path, reason="tesseract_not_installed")

        try:
            img = PILImage.open(file_path)
        except Exception as exc:
            logger.warning("이미지 파일 열기 실패 (%s): %s", file_path, exc)
            return _empty_image_doc(file_path, reason="file_open_error")

        try:
            data = pytesseract.image_to_data(
                img,
                lang="kor",
                output_type=pytesseract.Output.DICT,
            )
        except Exception as exc:
            logger.warning("OCR 실패: %s", exc)
            return _empty_image_doc(file_path, reason="ocr_error")

        all_cells: list[CellData] = []
        raw_parts: list[str] = []

        n_boxes = len(data["level"])
        for i in range(n_boxes):
            text = str(data["text"][i]).strip()
            if not text:
                continue
            try:
                conf = int(data["conf"][i])
            except (ValueError, TypeError):
                conf = 0
            if conf <= 30:
                continue

            x = data["left"][i]
            y = data["top"][i]
            address = f"ocr:{x},{y}"

            raw_parts.append(text)
            all_cells.append(CellData(
                address=address,
                value=text,
                data_type="text",
                neighbors={"confidence": conf, "x": x, "y": y},
            ))

        return ParsedDocument(
            file_path=file_path,
            file_type="image",
            raw_text=" ".join(raw_parts),
            structure={"ocr_blocks": len(all_cells)},
            cells=all_cells,
            metadata={"lang": "kor"},
        )


def _empty_image_doc(file_path: str, reason: str = "") -> ParsedDocument:
    """OCR 불가 시 반환하는 빈 ParsedDocument."""
    return ParsedDocument(
        file_path=file_path,
        file_type="image",
        raw_text="",
        structure={},
        cells=[],
        metadata={"fallback": True, "reason": reason},
    )
