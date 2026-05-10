import pytest
from unittest.mock import patch, MagicMock
from doc_intelligence.web.snapshot import capture_window_snapshot


def test_capture_returns_base64_string():
    fake_img = MagicMock()
    import io
    buf = io.BytesIO()
    from PIL import Image
    Image.new("RGB", (1, 1), "white").save(buf, format="PNG")
    fake_img_bytes = buf.getvalue()

    with patch("doc_intelligence.web.snapshot.pyautogui") as mock_pyautogui:
        mock_screenshot = MagicMock()
        mock_screenshot.save = MagicMock(side_effect=lambda buf, **kw: buf.write(fake_img_bytes))
        mock_pyautogui.screenshot.return_value = mock_screenshot
        with patch("doc_intelligence.web.snapshot._get_window_rect", return_value=(0, 0, 100, 100)):
            result = capture_window_snapshot("test.xlsx")

    assert isinstance(result, str)
    assert len(result) > 0


def test_capture_returns_none_when_window_not_found():
    with patch("doc_intelligence.web.snapshot._get_window_rect", return_value=None):
        result = capture_window_snapshot("nonexistent.xlsx")
    assert result is None
