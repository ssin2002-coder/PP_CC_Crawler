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
    """Excel.Application COM에서 Workbook을 파싱한다."""

    def parse_from_com(self, com_app, doc_obj=None) -> ParsedDocument:
        wb = doc_obj if doc_obj is not None else com_app.ActiveWorkbook
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

            # 열 너비 / 행 높이 (시트당 1회)
            col_widths = {}
            for c in range(col_start, col_end + 1):
                try:
                    col_widths[c] = round(float(ws.Columns(c).ColumnWidth), 1)
                except Exception:
                    col_widths[c] = 8.0

            row_heights = {}
            for r in range(row_start, row_end + 1):
                try:
                    row_heights[r] = round(float(ws.Rows(r).RowHeight), 1)
                except Exception:
                    row_heights[r] = 15.0

            # 병합 영역 수집
            merge_ranges = []
            seen_merges = set()

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

                    # 병합 정보
                    merge_rowspan = 1
                    merge_colspan = 1
                    merge_hidden = False
                    if is_merged:
                        merge_cells.append(address)
                        try:
                            ma = cell.MergeArea
                            ma_r, ma_c = ma.Row, ma.Column
                            ma_rs, ma_cs = ma.Rows.Count, ma.Columns.Count
                            if r == ma_r and c == ma_c:
                                merge_rowspan = ma_rs
                                merge_colspan = ma_cs
                                merge_key = f"{sheet_name}!R{ma_r}C{ma_c}"
                                if merge_key not in seen_merges:
                                    seen_merges.add(merge_key)
                                    merge_ranges.append({
                                        "row": ma_r, "col": ma_c,
                                        "rowspan": ma_rs, "colspan": ma_cs,
                                    })
                            else:
                                merge_hidden = True
                        except Exception:
                            pass

                    # 배경색
                    bg_color = None
                    try:
                        color_long = cell.Interior.Color
                        if color_long is not None:
                            cl = int(color_long)
                            # 16777215 = 흰색(#ffffff), 0 이하 = 자동
                            if 0 < cl < 16777215:
                                bg_color = f"#{cl & 0xFF:02x}{(cl >> 8) & 0xFF:02x}{(cl >> 16) & 0xFF:02x}"
                    except Exception:
                        pass

                    # 정렬
                    align = "general"
                    try:
                        ha = int(cell.HorizontalAlignment)
                        align = {-4131: "left", -4108: "center", -4152: "right"}.get(ha, "general")
                    except Exception:
                        pass

                    neighbors = {
                        "merged": is_merged,
                        "merge_hidden": merge_hidden,
                        "merge_rowspan": merge_rowspan,
                        "merge_colspan": merge_colspan,
                        "number_format": fmt,
                        "row": r,
                        "col": c,
                        "bg_color": bg_color,
                        "align": align,
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
                "col_widths": col_widths,
                "row_heights": row_heights,
                "merge_ranges": merge_ranges,
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
    """Word.Application COM에서 Document를 파싱한다."""

    def parse_from_com(self, com_app, doc_obj=None) -> ParsedDocument:
        doc = doc_obj if doc_obj is not None else com_app.ActiveDocument
        file_path = str(doc.FullName)

        all_cells: list[CellData] = []
        raw_parts: list[str] = []

        # 표 범위 수집 (문단 중복 제거용)
        table_ranges = []
        for table in doc.Tables:
            try:
                table_ranges.append((table.Range.Start, table.Range.End))
            except Exception:
                pass

        # Paragraphs 순회 — 표 내부 문단은 제외
        for i, para in enumerate(doc.Paragraphs):
            try:
                p_start = para.Range.Start
                in_table = any(ts <= p_start <= te for ts, te in table_ranges)
                if in_table:
                    continue
            except Exception:
                pass
            text = str(para.Range.Text).replace("\x07", "").replace("\r", "").strip()
            if text:
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
                    text = str(cell.Range.Text).replace("\x07", "").replace("\r", "").strip()
                    if text:
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
    """PowerPoint.Application COM에서 Presentation을 파싱한다."""

    def parse_from_com(self, com_app, doc_obj=None) -> ParsedDocument:
        prs = doc_obj if doc_obj is not None else com_app.ActivePresentation
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

    def parse_from_com(self, com_app, pd_doc=None) -> ParsedDocument:
        try:
            return self._parse_acrobat(com_app, pd_doc=pd_doc)
        except Exception as exc:
            logger.warning("Acrobat COM 파싱 실패, fallback 전환: %s", exc)
            return self._fallback_ocr(com_app)

    def _parse_acrobat(self, com_app, pd_doc=None) -> ParsedDocument:
        if pd_doc is not None:
            doc = pd_doc
        else:
            av_doc = com_app.GetActiveDoc()
            doc = av_doc.GetPDDoc()

        js = doc.GetJSObject()
        page_count = doc.GetNumPages()

        all_cells: list[CellData] = []
        raw_parts: list[str] = []

        for page_idx in range(page_count):
            for word_idx in range(js.getPageNumWords(page_idx)):
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

    @staticmethod
    def parse_from_file(file_path: str) -> ParsedDocument:
        """pypdf로 PDF 파일을 직접 파싱한다 (Acrobat COM 불필요)."""
        try:
            from pypdf import PdfReader
        except ImportError:
            logger.warning("pypdf 미설치 — PDF 파일 파싱 불가")
            return ParsedDocument(
                file_path=file_path, file_type="pdf", raw_text="",
                structure={}, cells=[],
                metadata={"fallback": True, "reason": "pypdf_not_installed"},
            )

        try:
            reader = PdfReader(file_path)
        except Exception as exc:
            logger.warning("PDF 파일 열기 실패 (%s): %s", file_path, exc)
            return ParsedDocument(
                file_path=file_path, file_type="pdf", raw_text="",
                structure={}, cells=[],
                metadata={"fallback": True, "reason": "file_open_error"},
            )

        all_cells: list[CellData] = []
        raw_parts: list[str] = []

        for page_idx, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            for line_idx, line in enumerate(text.split("\n")):
                line = line.strip()
                if line:
                    raw_parts.append(line)
                    address = f"pdf:p{page_idx}L{line_idx}"
                    all_cells.append(CellData(
                        address=address,
                        value=line,
                        data_type="text",
                        neighbors={"page": page_idx, "line": line_idx},
                    ))

        return ParsedDocument(
            file_path=file_path,
            file_type="pdf",
            raw_text="\n".join(raw_parts),
            structure={"page_count": len(reader.pages)},
            cells=all_cells,
            metadata={},
        )


# ──────────────────────────────────────────────
# ImageParser
# ──────────────────────────────────────────────

try:
    import asyncio as _asyncio
    from winrt.windows.media.ocr import OcrEngine as _OcrEngine
    from winrt.windows.graphics.imaging import BitmapDecoder as _BitmapDecoder
    from winrt.windows.storage import StorageFile as _StorageFile
    from winrt.windows.globalization import Language as _Language
    _WINRT_OCR_AVAILABLE = True
except ImportError:
    _WINRT_OCR_AVAILABLE = False


def _windows_ocr(file_path: str) -> list:
    """Windows 10/11 내장 OCR (Python winrt). Tesseract 불필요.

    반환: [{"text": str, "words": [{"text", "x", "y", "w", "h"}]}, ...]
    """
    if not _WINRT_OCR_AVAILABLE:
        logger.warning("winrt OCR 패키지 미설치")
        return []

    import os
    abs_path = os.path.abspath(file_path)

    async def _run():
        file = await _StorageFile.get_file_from_path_async(abs_path)
        stream = await file.open_read_async()
        decoder = await _BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()
        engine = _OcrEngine.try_create_from_language(_Language("ko"))
        if engine is None:
            engine = _OcrEngine.try_create_from_user_profile_languages()
        result = await engine.recognize_async(bitmap)
        lines = []
        for line in result.lines:
            words = []
            for word in line.words:
                r = word.bounding_rect
                words.append({
                    "text": word.text,
                    "x": int(r.x), "y": int(r.y),
                    "w": int(r.width), "h": int(r.height),
                })
            lines.append({"text": line.text, "words": words})
        return lines

    try:
        return _asyncio.run(_run())
    except Exception as exc:
        logger.warning("Windows OCR 실패: %s", exc)
        return []


class ImageParser(BaseParser):
    """이미지 파일 OCR 파서.

    com_app 인자로 이미지 파일 경로(str)를 받는다.
    우선순위: pytesseract → Windows 내장 OCR → graceful fail.
    """

    def parse_from_com(self, com_app) -> ParsedDocument:
        file_path = str(com_app) if com_app else ""

        if not _TESSERACT_AVAILABLE:
            logger.info("pytesseract 미설치 — Windows 내장 OCR 시도")
            return self._windows_ocr_parse(file_path)

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

    @staticmethod
    def _ocr_postprocess(text: str) -> str:
        """OCR 오인식 보정 — 기술 문서 특수문자/단위."""
        import re
        corrections = {
            "두꼐": "두께", "두게": "두께",
            "Q/sq": "Ω/sq", "0/sq": "Ω/sq",
            "니m": "μm", "um": "μm", "u m": "μm",
            "20.5": "≥0.5", "20.5u": "≥0.5μ",
            ">=": "≥", "<=": "≤",
            "ea/m2": "ea/m²", "m2": "m²",
            "검사제": "검사자:",
            "0/0": "%",
        }
        for wrong, right in corrections.items():
            text = text.replace(wrong, right)
        return text

    @staticmethod
    def _build_ocr_grid(lines: list) -> tuple:
        """OCR 라인을 y행 클러스터 + 헤더 기반 열 경계 매핑으로 구조화.

        반환: (grid_rows: list[list[str]], max_cols: int)
        """
        all_words = []
        for line_data in lines:
            for word in line_data.get("words", []):
                all_words.append(word)
        if not all_words:
            return [], 0

        # 1단계: y좌표 행 클러스터링 (허용 12px)
        all_words.sort(key=lambda w: (w.get("y", 0), w.get("x", 0)))
        raw_rows = []
        curr = [all_words[0]]
        cy = all_words[0].get("y", 0)
        for w in all_words[1:]:
            if abs(w.get("y", 0) - cy) <= 12:
                curr.append(w)
            else:
                curr.sort(key=lambda w: w.get("x", 0))
                raw_rows.append(curr)
                curr = [w]
                cy = w.get("y", 0)
        if curr:
            curr.sort(key=lambda w: w.get("x", 0))
            raw_rows.append(curr)

        # 2단계: 헤더 행 자동 감지 — 가장 넓게 퍼진 행 (x범위 최대)
        best_row = None
        best_spread = 0
        for row in raw_rows:
            if len(row) < 5:
                continue
            xs = [w.get("x", 0) for w in row]
            spread = max(xs) - min(xs)
            if spread > best_spread:
                best_spread = spread
                best_row = row

        if best_row is None:
            grid_rows = [[" ".join(w.get("text", "") for w in row)] for row in raw_rows]
            return grid_rows, 1

        # 헤더 워드의 x 중심을 열 앵커로 사용
        col_anchors = []
        for w in best_row:
            cx = w.get("x", 0) + w.get("w", 0) // 2
            col_anchors.append(cx)
        num_cols = len(col_anchors)

        def find_col(x):
            """x좌표에 가장 가까운 열 앵커 찾기"""
            min_dist = float("inf")
            best = 0
            for i, anchor in enumerate(col_anchors):
                d = abs(x - anchor)
                if d < min_dist:
                    min_dist = d
                    best = i
            return best

        # 3단계: 모든 행을 열 앵커에 매핑
        grid_rows = []
        for row in raw_rows:
            cells = [""] * num_cols
            for w in row:
                wx = w.get("x", 0) + w.get("w", 0) // 2
                ci = find_col(wx)
                if cells[ci]:
                    cells[ci] += " " + w.get("text", "")
                else:
                    cells[ci] = w.get("text", "")

            # 후처리: col0에 "숫자 텍스트"가 합쳐진 경우 분리 → col0=숫자, col1=텍스트
            import re
            if num_cols >= 2 and cells[0] and not cells[1]:
                m = re.match(r"^(\d+)\s+(.+)$", cells[0].strip())
                if m:
                    cells[0] = m.group(1)
                    cells[1] = m.group(2)

            grid_rows.append(cells)

        return grid_rows, num_cols

    def _windows_ocr_parse(self, file_path: str) -> ParsedDocument:
        """Windows 내장 OCR로 이미지를 파싱한다."""
        import os
        if not os.path.isfile(file_path):
            return _empty_image_doc(file_path, reason="file_open_error")

        lines = _windows_ocr(file_path)
        if not lines:
            return _empty_image_doc(file_path, reason="windows_ocr_no_result")

        all_cells: list[CellData] = []
        raw_parts: list[str] = []

        grid_rows, max_cols = self._build_ocr_grid(lines)

        for r_idx, row_cells in enumerate(grid_rows):
            for c_idx, cell_text in enumerate(row_cells):
                text = self._ocr_postprocess(cell_text.strip())
                if not text:
                    continue
                address = f"ocr_tbl:R{r_idx}C{c_idx}"
                raw_parts.append(text)
                all_cells.append(CellData(
                    address=address,
                    value=text,
                    data_type="text",
                    neighbors={"row": r_idx, "col": c_idx, "ocr_engine": "windows"},
                ))

        return ParsedDocument(
            file_path=file_path,
            file_type="image",
            raw_text=" ".join(raw_parts),
            structure={"ocr_blocks": len(all_cells), "ocr_rows": len(grid_rows), "ocr_cols": max_cols},
            cells=all_cells,
            metadata={"lang": "ko", "ocr_engine": "windows"},
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
