# doc_intelligence/web/snapshot.py
"""윈도우 캡처 — win32gui + win32ui 기반 스냅샷 (pyautogui 미사용)"""
import base64
import ctypes
import io
import logging

logger = logging.getLogger(__name__)

# DPI 인식 설정 (프로세스 전역, 1회만 호출)
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

try:
    import win32gui
    import win32ui
    import win32con
    _WIN32_AVAILABLE = True
except ImportError:
    win32gui = None
    win32ui = None
    win32con = None
    _WIN32_AVAILABLE = False

try:
    from PIL import Image as PILImage
    _PIL_AVAILABLE = True
except ImportError:
    PILImage = None
    _PIL_AVAILABLE = False


def _get_window_rect(filename: str):
    if not _WIN32_AVAILABLE:
        return None
    import os
    # 확장자 제거 — "파일.xlsx" → "파일" (윈도우 타이틀은 확장자 없이 표시)
    name_no_ext = os.path.splitext(filename)[0] if filename else filename
    result = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if name_no_ext and name_no_ext in title:
                rect = win32gui.GetWindowRect(hwnd)
                result.append(rect)
    try:
        win32gui.EnumWindows(callback, None)
    except Exception:
        pass
    return result[0] if result else None


def capture_window_snapshot(filename: str) -> str | None:
    if not _WIN32_AVAILABLE or not _PIL_AVAILABLE:
        return None
    rect = _get_window_rect(filename)
    if rect is None:
        return None
    left, top, right, bottom = rect
    width = right - left
    height = bottom - top
    if width <= 0 or height <= 0:
        return None
    try:
        hwnd_desktop = win32gui.GetDesktopWindow()
        dc_desktop = win32gui.GetWindowDC(hwnd_desktop)
        dc_obj = win32ui.CreateDCFromHandle(dc_desktop)
        mem_dc = dc_obj.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(dc_obj, width, height)
        mem_dc.SelectObject(bitmap)
        mem_dc.BitBlt((0, 0), (width, height), dc_obj, (left, top), win32con.SRCCOPY)

        bmp_info = bitmap.GetInfo()
        bmp_bits = bitmap.GetBitmapBits(True)
        img = PILImage.frombuffer(
            "RGB", (bmp_info["bmWidth"], bmp_info["bmHeight"]),
            bmp_bits, "raw", "BGRX", 0, 1,
        )

        mem_dc.DeleteDC()
        dc_obj.DeleteDC()
        win32gui.ReleaseDC(hwnd_desktop, dc_desktop)
        win32gui.DeleteObject(bitmap.GetHandle())

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception as e:
        logger.warning("스냅샷 캡처 실패: %s", e)
        return None
