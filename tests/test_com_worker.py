"""
test_com_worker.py
COM 없는 환경에서도 동작하는 ComWorker 테스트.
"""
import pytest
from unittest.mock import MagicMock, patch


# ComWorker import
from doc_intelligence.com_worker import ComWorker


class TestComWorkerRetryOnFailure:
    """test_com_worker_retry_on_failure — 3회 시도 중 3번째 성공."""

    def test_succeeds_on_third_attempt(self):
        worker = ComWorker(max_retries=3, timeout=10)

        call_count = 0

        def flaky_func():
            nonlocal call_count
            call_count += 1
            if call_count < 3:
                raise RuntimeError("일시적 오류")
            return "success"

        with patch("time.sleep"):  # sleep 스킵
            result = worker.execute(flaky_func)

        assert result == "success"
        assert call_count == 3


class TestComWorkerMaxRetriesExceeded:
    """test_com_worker_max_retries_exceeded — max_retries 초과 시 예외."""

    def test_raises_after_max_retries(self):
        worker = ComWorker(max_retries=3, timeout=10)

        def always_fails():
            raise ValueError("항상 실패")

        with patch("time.sleep"):
            with pytest.raises(ValueError, match="항상 실패"):
                worker.execute(always_fails)

    def test_sleep_called_between_retries(self):
        """재시도 사이에 sleep이 (max_retries - 1)회 호출되는지 확인."""
        worker = ComWorker(max_retries=3, timeout=10)

        def always_fails():
            raise RuntimeError("fail")

        with patch("time.sleep") as mock_sleep:
            with pytest.raises(RuntimeError):
                worker.execute(always_fails)

        # 3회 시도, 마지막엔 sleep 없음 → 2회
        assert mock_sleep.call_count == 2


class TestComWorkerGetActiveApp:
    """test_com_worker_get_active_app — COM 없는 환경에서 None 반환."""

    def test_returns_none_when_com_unavailable(self):
        worker = ComWorker()
        result = worker.get_active_app("Excel.Application")
        # COM 없거나 앱 미실행 시 None
        assert result is None

    def test_returns_none_for_any_prog_id(self):
        worker = ComWorker()
        for prog_id in ["Excel.Application", "Word.Application", "PowerPoint.Application"]:
            result = worker.get_active_app(prog_id)
            assert result is None, f"{prog_id} should return None"


class TestDetectOpenDocuments:
    """test_detect_open_documents — COM 없을 때 빈 리스트 반환."""

    def test_returns_empty_list_when_com_unavailable(self):
        worker = ComWorker()
        result = worker.detect_open_documents()
        assert isinstance(result, list)
        assert result == []

    def test_returns_list_type(self):
        worker = ComWorker()
        result = worker.detect_open_documents()
        assert isinstance(result, list)


class TestComSession:
    """com_session 컨텍스트 매니저 — COM 없는 환경에서도 예외 없이 진입/탈출."""

    def test_context_manager_no_error_without_com(self):
        worker = ComWorker()
        entered = False
        exited = False

        with worker.com_session():
            entered = True
        exited = True

        assert entered
        assert exited

    def test_context_manager_yields_none(self):
        worker = ComWorker()
        with worker.com_session() as val:
            assert val is None
