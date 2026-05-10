"""
com_worker.py
COM 프로세스 격리 래퍼 — STA 세션, 재시도 3회, 타임아웃.

COM 미설치(pythoncom/win32com 없음) 환경에서도 import 에러 없이 동작.
"""
import time
from contextlib import contextmanager

# COM 라이브러리 조건부 import
try:
    import pythoncom
    import win32com.client
    _COM_AVAILABLE = True
except ImportError:
    pythoncom = None
    win32com = None
    _COM_AVAILABLE = False


class ComWorker:
    """COM 호출을 안정화하는 래퍼.

    - STA 모드 COM 세션 관리 (com_session)
    - 재시도 정책 (execute)
    - 실행 중인 COM 앱 감지 (get_active_app, detect_open_documents)
    """

    def __init__(self, max_retries: int = 3, timeout: int = 10):
        self.max_retries = max_retries
        self.timeout = timeout

    def execute(self, func, *args, **kwargs):
        """COM 함수를 재시도 정책으로 실행. 실패 시 1초 대기 후 재시도.

        Args:
            func: 실행할 callable.
            *args, **kwargs: func에 전달할 인자.

        Returns:
            func의 반환값.

        Raises:
            마지막 시도에서 발생한 예외.
        """
        last_error = None
        for attempt in range(1, self.max_retries + 1):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                last_error = e
                if attempt < self.max_retries:
                    time.sleep(1)
        raise last_error

    def get_active_app(self, prog_id: str):
        """실행 중인 COM 앱에 연결. 없으면 None.

        Args:
            prog_id: COM ProgID (예: "Excel.Application").

        Returns:
            COM 앱 객체 또는 None.
        """
        if not _COM_AVAILABLE:
            return None
        try:
            return win32com.client.GetActiveObject(prog_id)
        except Exception:
            try:
                return win32com.client.Dispatch(prog_id)
            except Exception:
                return None

    def detect_open_documents(self) -> list:
        """Excel, Word, PPT에서 열린 문서 목록 감지.

        Returns:
            열린 문서 정보 dict 리스트.
            각 항목: {"app": prog_id, "name": 문서명, "path": 전체경로}
        """
        results = []
        if not _COM_AVAILABLE:
            return results

        app_configs = [
            ("Excel.Application", "Workbooks"),
            ("Word.Application", "Documents"),
            ("PowerPoint.Application", "Presentations"),
        ]

        for prog_id, collection_name in app_configs:
            app = self.get_active_app(prog_id)
            if app is None:
                continue
            try:
                collection = getattr(app, collection_name)
                for i in range(1, collection.Count + 1):
                    try:
                        doc = collection.Item(i)
                        results.append({
                            "app": prog_id,
                            "app_obj": app,
                            "name": getattr(doc, "Name", ""),
                            "path": getattr(doc, "FullName", ""),
                        })
                    except Exception:
                        continue
            except Exception:
                continue

        return results

    @contextmanager
    def com_session(self):
        """STA COM 세션 컨텍스트 매니저.

        COM 사용 가능 시 CoInitialize/CoUninitialize로 STA 모드 보장.
        COM 미설치 환경에서도 예외 없이 진입/탈출.

        Yields:
            None
        """
        if _COM_AVAILABLE:
            pythoncom.CoInitialize()
        try:
            yield None
        finally:
            if _COM_AVAILABLE:
                pythoncom.CoUninitialize()
