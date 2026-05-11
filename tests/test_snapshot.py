import pytest
from unittest.mock import patch, MagicMock
from doc_intelligence.web.snapshot import capture_window_snapshot


def test_capture_returns_base64_string():
    import io
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (100, 100), "white").save(buf, format="PNG")
    fake_img_bytes = buf.getvalue()

    fake_bitmap = MagicMock()
    fake_bitmap.GetInfo.return_value = {"bmWidth": 100, "bmHeight": 100}
    fake_bitmap.GetBitmapBits.return_value = b"\x00\x00\xff\x00" * (100 * 100)
    fake_bitmap.GetHandle.return_value = 0

    with patch("doc_intelligence.web.snapshot._WIN32_AVAILABLE", True), \
         patch("doc_intelligence.web.snapshot._PIL_AVAILABLE", True), \
         patch("doc_intelligence.web.snapshot._get_window_rect", return_value=(0, 0, 100, 100)), \
         patch("doc_intelligence.web.snapshot.win32gui") as mock_gui, \
         patch("doc_intelligence.web.snapshot.win32ui") as mock_ui, \
         patch("doc_intelligence.web.snapshot.PILImage") as mock_pil:
        mock_gui.GetDesktopWindow.return_value = 0
        mock_gui.GetWindowDC.return_value = 0
        mock_dc = MagicMock()
        mock_mem_dc = MagicMock()
        mock_ui.CreateDCFromHandle.return_value = mock_dc
        mock_dc.CreateCompatibleDC.return_value = mock_mem_dc
        mock_ui.CreateBitmap.return_value = fake_bitmap

        mock_img = MagicMock()
        mock_img.save = MagicMock(side_effect=lambda buf, **kw: buf.write(fake_img_bytes))
        mock_pil.frombuffer.return_value = mock_img

        result = capture_window_snapshot("test.xlsx")

    assert isinstance(result, str)
    assert len(result) > 0


def test_capture_returns_none_when_window_not_found():
    with patch("doc_intelligence.web.snapshot._get_window_rect", return_value=None):
        result = capture_window_snapshot("nonexistent.xlsx")
    assert result is None
