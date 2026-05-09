"""
region_linker.py — 드래그 영역 연결 + DPI 스케일링
드래그로 선택된 화면 영역을 앱별 문서 좌표(셀, 단락 등)와 연결하는 모듈
"""
import logging
from dataclasses import dataclass
from typing import Optional

try:
    import win32gui
    import psutil
    _WIN32_AVAILABLE = True
except ImportError:
    _WIN32_AVAILABLE = False

try:
    import pyautogui
    _PYAUTOGUI_AVAILABLE = True
except ImportError:
    _PYAUTOGUI_AVAILABLE = False

logger = logging.getLogger(__name__)


@dataclass
class LinkedRegion:
    app_name: str       # "EXCEL.EXE" 등
    doc_name: str
    location: str       # "Sheet1!B4", "para:3" 등
    screen_rect: tuple  # (x, y, w, h)
    screenshot: Optional[bytes]


class RegionLinker:
    def __init__(self, storage=None):
        self.storage = storage
        self.current_regions = []

    # ──────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────

    def create_rule(self, regions, rule_type, rule_name):
        """LinkedRegion 리스트로 룰 데이터 생성 후 반환.

        storage가 있으면 storage.save_rule을 호출하여 DB에 저장한다.
        반환값: {"name", "rule_type", "regions", "params", "id"} 딕셔너리
        """
        region_data = []
        for r in regions:
            region_data.append({
                "app_name": r.app_name,
                "doc_name": r.doc_name,
                "location": r.location,
                "screen_rect": list(r.screen_rect),
            })

        params = {"region_count": len(regions)}
        rule = {
            "name": rule_name,
            "rule_type": rule_type,
            "regions": region_data,
            "params": params,
            "id": None,
        }

        if self.storage is not None:
            try:
                rule_id = self.storage.save_rule(
                    name=rule_name,
                    rule_type=rule_type,
                    conditions={"regions": region_data},
                    actions=params,
                )
                rule["id"] = rule_id
            except Exception as e:
                logger.warning("storage.save_rule 실패: %s", e)

        return rule

    # ──────────────────────────────────────────────
    # DPI 유틸
    # ──────────────────────────────────────────────

    def _apply_dpi_scale(self, physical_coord, scale_factor):
        """물리 좌표 → 논리 좌표 변환 (physical_coord / scale_factor, 정수 반환)"""
        return int(physical_coord / scale_factor)

    def _get_dpi_scale(self):
        """시스템 DPI 배율 반환. ctypes 획득 실패 시 1.0"""
        try:
            import ctypes
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return dpi / 96.0
        except Exception as e:
            logger.debug("DPI 획득 실패, 기본값 1.0 사용: %s", e)
            return 1.0

    # ──────────────────────────────────────────────
    # 앱 식별
    # ──────────────────────────────────────────────

    def _get_app_from_point(self, x, y):
        """좌표(x, y)에 해당하는 앱 프로세스 이름 반환.

        win32gui.WindowFromPoint + psutil로 획득. 실패 시 None.
        """
        if not _WIN32_AVAILABLE:
            logger.debug("win32gui/psutil 사용 불가, None 반환")
            return None
        try:
            hwnd = win32gui.WindowFromPoint((x, y))
            if hwnd == 0:
                return None
            import win32process
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            return proc.name().upper()
        except Exception as e:
            logger.debug("앱 식별 실패: %s", e)
            return None

    # ──────────────────────────────────────────────
    # 앱별 좌표 변환
    # ──────────────────────────────────────────────

    def _screen_to_excel_cell(self, excel_app, x, y):
        """화면 좌표 → Excel 셀 주소 변환.

        excel_app.ActiveWindow.RangeFromPoint(x, y) 사용.
        실패 시 "screen:{x},{y}" 반환.
        """
        try:
            rng = excel_app.ActiveWindow.RangeFromPoint(x, y)
            return rng.Address
        except Exception as e:
            logger.debug("Excel RangeFromPoint 실패: %s", e)
            return f"screen:{x},{y}"

    def _screen_to_word_location(self, word_app, x, y):
        """화면 좌표 → Word 위치 문자열 반환.

        Word에는 RangeFromPoint API가 없으므로 "screen:{x},{y}" 반환.
        """
        return f"screen:{x},{y}"

    def _screen_to_ppt_location(self, ppt_app, x, y):
        """화면 좌표 → PowerPoint 위치 문자열 반환."""
        return f"screen:{x},{y}"

    def _screen_to_pdf_location(self, acrobat_app, x, y):
        """화면 좌표 → PDF 위치 문자열 반환.

        AcroExch COM 객체로 페이지 번호 획득 시도. 실패 시 "screen:{x},{y}".
        """
        try:
            page_num = acrobat_app.GetPageNumFromPoint(x, y)
            return f"page:{page_num}"
        except Exception as e:
            logger.debug("PDF 페이지 번호 획득 실패: %s", e)
            return f"screen:{x},{y}"

    # ──────────────────────────────────────────────
    # 화면 캡처
    # ──────────────────────────────────────────────

    def capture_region(self, rect):
        """지정한 rect=(x, y, w, h) 영역 스크린샷을 bytes로 반환.

        pyautogui 사용. 실패 시 None.
        """
        if not _PYAUTOGUI_AVAILABLE:
            logger.debug("pyautogui 사용 불가, None 반환")
            return None
        try:
            import io
            img = pyautogui.screenshot(region=rect)
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            return buf.getvalue()
        except Exception as e:
            logger.debug("스크린샷 캡처 실패: %s", e)
            return None
