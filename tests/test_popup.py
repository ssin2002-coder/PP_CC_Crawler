# tests/test_popup.py
"""ParseResultPopup 동작 검증.

주의: tkinter 를 실제로 띄우지 않고 _root 를 MagicMock 으로 대체해
add_records → _maybe_refresh 가드, <Map> 이벤트 처리 로직만 단위
테스트한다. GUI 통합 테스트가 아니다."""
from unittest.mock import MagicMock
from word_crawler import ParseResultPopup


def _make_popup():
    p = ParseResultPopup(on_save_all=None, db_path=None)
    p._alive = True
    p._root = MagicMock()
    p._refresh_tree = MagicMock()
    return p


class TestAddRecordsDeferRefresh:
    """창이 최소화/숨김 상태일 때 _refresh_tree 가 호출되지 않아야 함."""

    def test_iconic_does_not_refresh(self):
        p = _make_popup()
        p._root.state.return_value = 'iconic'
        p.add_records('a.docx', '2024-05-03', [{'k': 'v'}], 'h1')
        # add_records 는 메인 스레드에 _maybe_refresh 위임만 함
        p._root.after.assert_called_once()
        cb = p._root.after.call_args[0][1]
        # 콜백 직접 실행해도 iconic 이면 _refresh_tree 호출 안 함
        cb()
        p._refresh_tree.assert_not_called()
        # 데이터는 누적됨, dirty 플래그 유지
        assert p._dirty is True
        assert len(p._pending) == 1

    def test_withdrawn_does_not_refresh(self):
        p = _make_popup()
        p._root.state.return_value = 'withdrawn'
        p.add_records('a.docx', '2024-05-03', [{'k': 'v'}], 'h1')
        p._root.after.call_args[0][1]()
        p._refresh_tree.assert_not_called()
        assert p._dirty is True

    def test_normal_refreshes(self):
        p = _make_popup()
        p._root.state.return_value = 'normal'
        p.add_records('a.docx', '2024-05-03', [{'k': 'v'}], 'h1')
        cb = p._root.after.call_args[0][1]
        cb()
        p._refresh_tree.assert_called_once()
        assert p._dirty is False

    def test_does_not_call_state_directly(self):
        """root.state() 가 add_records 본체(=백그라운드 스레드 가능) 에서
        호출되면 안 되고, 반드시 after() 로 위임된 콜백 안에서만 호출."""
        p = _make_popup()
        p._root.state.return_value = 'iconic'
        p.add_records('a.docx', '2024-05-03', [{'k': 'v'}], 'h1')
        # add_records 자체는 state() 를 호출하지 않음 — 위임만 함
        p._root.state.assert_not_called()
        # 콜백 실행 후에야 state() 호출
        cb = p._root.after.call_args[0][1]
        cb()
        p._root.state.assert_called_once()

    def test_accumulates_multiple_pending_while_iconic(self):
        p = _make_popup()
        p._root.state.return_value = 'iconic'
        for i in range(3):
            p.add_records(f'a{i}.docx', '2024-05-03', [{'k': i}], f'h{i}')
        # 3 번 모두 누적되고 한 번도 갱신 안 됨
        assert len(p._pending) == 3
        assert p._dirty is True
        # 각 호출의 콜백을 실행해도 여전히 갱신 안 됨
        for call in p._root.after.call_args_list:
            call[0][1]()
        p._refresh_tree.assert_not_called()


class TestOnWindowMap:
    """창이 다시 visible 해지면(<Map> 이벤트) 누적된 dirty 가 있을 때만 갱신."""

    def test_refreshes_when_dirty(self):
        p = _make_popup()
        p._dirty = True
        event = MagicMock()
        event.widget = p._root  # is 비교를 위한 동일 객체
        p._on_window_map(event)
        p._refresh_tree.assert_called_once()
        assert p._dirty is False

    def test_no_refresh_when_clean(self):
        p = _make_popup()
        p._dirty = False
        event = MagicMock()
        event.widget = p._root
        p._on_window_map(event)
        p._refresh_tree.assert_not_called()

    def test_ignores_child_widget_events(self):
        # 자식 위젯의 <Map> 이벤트는 무시 (root 의 것만 처리)
        p = _make_popup()
        p._dirty = True
        event = MagicMock()
        event.widget = MagicMock()  # root 와 다른 객체
        p._on_window_map(event)
        p._refresh_tree.assert_not_called()
        assert p._dirty is True  # dirty 유지


class TestRestoreWindow:
    def test_iconic_deiconifies(self):
        p = _make_popup()
        p._root.state.return_value = 'iconic'
        p._restore_window()
        p._root.deiconify.assert_called_once()
        p._root.lift.assert_called_once()

    def test_withdrawn_deiconifies(self):
        p = _make_popup()
        p._root.state.return_value = 'withdrawn'
        p._restore_window()
        p._root.deiconify.assert_called_once()
