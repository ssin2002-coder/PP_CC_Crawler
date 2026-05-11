import pytest
from unittest.mock import patch, MagicMock
from doc_intelligence.web.snapshot import capture_window_snapshot


def test_capture_returns_base64_string():
    fake_raw = MagicMock()
    fake_raw.size = (100, 100)
    fake_raw.rgb = b"\x00" * (100 * 100 * 3)

    mock_sct = MagicMock()
    mock_sct.grab.return_value = fake_raw
    mock_sct.__enter__ = MagicMock(return_value=mock_sct)
    mock_sct.__exit__ = MagicMock(return_value=False)

    with patch("doc_intelligence.web.snapshot.mss") as mock_mss_mod:
        mock_mss_mod.mss.return_value = mock_sct
        with patch("doc_intelligence.web.snapshot._get_window_rect", return_value=(0, 0, 100, 100)):
            result = capture_window_snapshot("test.xlsx")

    assert isinstance(result, str)
    assert len(result) > 0


def test_capture_returns_none_when_window_not_found():
    with patch("doc_intelligence.web.snapshot._get_window_rect", return_value=None):
        result = capture_window_snapshot("nonexistent.xlsx")
    assert result is None
