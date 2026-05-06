# tests/test_popup.py
"""ParseResultPopup 동작 검증.

새 architecture(메인 스레드 root + Toplevel 한 번 빌드 후 deiconify/withdraw
토글, race/alive 가드 제거) 의 핵심 불변성을 단위 테스트한다. tkinter 를
실제로 띄우지 않고 root/window 를 MagicMock 으로 대체."""
from unittest.mock import MagicMock, patch
from word_crawler import ParseResultPopup


def _make_popup_with_built_ui():
    """UI 가 이미 빌드된 상태(=_window 가 set 됨) 의 popup 시뮬레이션."""
    root = MagicMock()
    p = ParseResultPopup(root)
    p._window = MagicMock()
    p._refresh_tree = MagicMock()
    p._build_ui = MagicMock()  # 다시 호출되면 안 됨
    return p


class TestAddRecordsDedup:
    """같은 (doc_name, date_str) 재파싱 시 _pending 이 누적되지 않고 교체됨."""

    def test_same_doc_and_date_replaces(self):
        p = _make_popup_with_built_ui()
        p._window.state.return_value = 'normal'
        p._do_add_records('a.docx', '2024-05-03', [{'k': 1}], 'h1')
        p._do_add_records('a.docx', '2024-05-03', [{'k': 2}], 'h2')
        # 같은 (doc, date) 이므로 한 개만 유지, 최신 records/hash
        assert len(p._pending) == 1
        assert p._pending[0] == ('a.docx', '2024-05-03', [{'k': 2}], 'h2')

    def test_different_doc_keeps_both(self):
        p = _make_popup_with_built_ui()
        p._window.state.return_value = 'normal'
        p._do_add_records('a.docx', '2024-05-03', [{'k': 1}], 'h1')
        p._do_add_records('b.docx', '2024-05-03', [{'k': 2}], 'h2')
        assert len(p._pending) == 2

    def test_same_doc_different_date_keeps_both(self):
        p = _make_popup_with_built_ui()
        p._window.state.return_value = 'normal'
        p._do_add_records('a.docx', '2024-05-03', [{'k': 1}], 'h1')
        p._do_add_records('a.docx', '2024-05-04', [{'k': 2}], 'h2')
        assert len(p._pending) == 2

    def test_three_calls_same_key_yields_one(self):
        p = _make_popup_with_built_ui()
        p._window.state.return_value = 'normal'
        for i in range(3):
            p._do_add_records('a.docx', '2024-05-03', [{'k': i}], f'h{i}')
        assert len(p._pending) == 1
        assert p._pending[0][3] == 'h2'


class TestAddRecordsRouting:
    """add_records 는 절대 위젯을 직접 만지지 않고 root.after 로 위임."""

    def test_add_records_uses_after(self):
        root = MagicMock()
        p = ParseResultPopup(root)
        p.add_records('a.docx', '2024-05-03', [{'k': 1}], 'h1')
        root.after.assert_called_once()
        args, _ = root.after.call_args
        assert args[0] == 0
        # 콜백 실행 시 _do_add_records 가 호출되어야 함
        p._do_add_records = MagicMock()
        callback = args[1]
        callback()
        p._do_add_records.assert_called_once_with(
            'a.docx', '2024-05-03', [{'k': 1}], 'h1')

    def test_show_window_uses_after(self):
        root = MagicMock()
        p = ParseResultPopup(root)
        p.show_window()
        root.after.assert_called_once_with(0, p._do_show_window)


class TestVisibilityGating:
    """창이 minimized/withdrawn 일 때 _refresh_tree 호출되지 않고 _dirty 만 True."""

    def test_iconic_skips_refresh(self):
        p = _make_popup_with_built_ui()
        p._window.state.return_value = 'iconic'
        p._do_add_records('a.docx', '2024-05-03', [{'k': 1}], 'h1')
        p._refresh_tree.assert_not_called()
        assert p._dirty is True
        assert len(p._pending) == 1

    def test_withdrawn_skips_refresh(self):
        p = _make_popup_with_built_ui()
        p._window.state.return_value = 'withdrawn'
        p._do_add_records('a.docx', '2024-05-03', [{'k': 1}], 'h1')
        p._refresh_tree.assert_not_called()
        assert p._dirty is True

    def test_normal_refreshes(self):
        p = _make_popup_with_built_ui()
        p._window.state.return_value = 'normal'
        p._do_add_records('a.docx', '2024-05-03', [{'k': 1}], 'h1')
        p._refresh_tree.assert_called_once()


class TestShowWindow:
    """show_window 는 deiconify + lift + focus + 누적 dirty 갱신."""

    def test_deiconifies_existing_window(self):
        p = _make_popup_with_built_ui()
        p._dirty = True
        p._do_show_window()
        p._window.deiconify.assert_called_once()
        p._window.lift.assert_called_once()
        p._window.focus_force.assert_called_once()
        # dirty 였으므로 _refresh_tree 호출 + dirty 해제
        p._refresh_tree.assert_called_once()
        assert p._dirty is False

    def test_no_refresh_when_clean(self):
        p = _make_popup_with_built_ui()
        p._dirty = False
        p._do_show_window()
        p._window.deiconify.assert_called_once()
        p._refresh_tree.assert_not_called()

    def test_builds_ui_if_window_missing(self):
        root = MagicMock()
        p = ParseResultPopup(root)

        def fake_build():
            p._window = MagicMock()
            p._window.state.return_value = 'normal'
        p._build_ui = MagicMock(side_effect=fake_build)
        p._refresh_tree = MagicMock()

        p._do_show_window()
        p._build_ui.assert_called_once()
        p._window.deiconify.assert_called_once()


class TestOnClose:
    """X 버튼은 destroy 가 아니라 withdraw — 트레이로 숨김 후 재사용 가능."""

    def test_close_withdraws_window(self):
        p = _make_popup_with_built_ui()
        p._on_close()
        p._window.withdraw.assert_called_once()
        p._window.destroy.assert_not_called()

    def test_close_no_window_is_noop(self):
        root = MagicMock()
        p = ParseResultPopup(root)
        # _window=None 인 상태에서 호출되어도 예외 없이 종료
        p._on_close()


class TestNoSelfTkInPopup:
    """ParseResultPopup 인스턴스가 직접 tk.Tk() 를 만들지 않는지 확인."""

    def test_init_does_not_call_tk_Tk(self):
        with patch('word_crawler.tk.Tk') as mock_tk:
            root = MagicMock()
            ParseResultPopup(root)
            mock_tk.assert_not_called()

    def test_add_records_does_not_call_tk_Tk(self):
        root = MagicMock()
        p = ParseResultPopup(root)
        with patch('word_crawler.tk.Tk') as mock_tk:
            p.add_records('a.docx', '2024-05-03', [{'k': 1}], 'h1')
            mock_tk.assert_not_called()
