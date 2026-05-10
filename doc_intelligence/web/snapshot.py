# doc_intelligence/web/snapshot.py
"""윈도우 캡처 — pyautogui 기반 스냅샷"""
import base64
import io
import logging

logger = logging.getLogger(__name__)

try:
    import pyautogui
    _PYAUTOGUI_AVAILABLE = True
except ImportError:
    pyautogui = None
    _PYAUTOGUI_AVAILABLE = False

try:
    import win32gui
    _WIN32GUI_AVAILABLE = True
except ImportError:
    win32gui = None
    _WIN32GUI_AVAILABLE = False


def _get_window_rect(filename: str):
    if not _WIN32GUI_AVAILABLE:
        return None
    result = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if filename in title:
                rect = win32gui.GetWindowRect(hwnd)
                result.append(rect)
    try:
        win32gui.EnumWindows(callback, None)
    except Exception:
        pass
    return result[0] if result else None


def capture_window_snapshot(filename: str) -> str | None:
    if pyautogui is None:
        return None
    rect = _get_window_rect(filename)
    if rect is None:
        return None
    left, top, right, bottom = rect
    width = right - left
    height = bottom - top
    try:
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        buf = io.BytesIO()
        screenshot.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception as e:
        logger.warning("스냅샷 캡처 실패: %s", e)
        return None
