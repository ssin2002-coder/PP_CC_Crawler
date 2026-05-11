"""
com_worker.py
COM 프로세스 격리 래퍼 — STA 세션, 재시도 3회, 타임아웃.

COM 미설치(pythoncom/win32com 없음) 환경에서도 import 에러 없이 동작.
"""
import os
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
                            "doc_obj": doc,
                            "name": getattr(doc, "Name", ""),
                            "path": getattr(doc, "FullName", ""),
                        })
                    except Exception:
                        continue
            except Exception:
                continue

        # Acrobat PDF 감지 — Dispatch 사용 안 함 (새 인스턴스 생성 방지)
        try:
            acrobat = None
            if _COM_AVAILABLE:
                try:
                    acrobat = win32com.client.GetActiveObject("AcroExch.App")
                except Exception:
                    pass
            if acrobat is not None:
                num_docs = acrobat.GetNumAVDocs()
                for i in range(num_docs):
                    try:
                        av_doc = acrobat.GetAVDoc(i)
                        pd_doc = av_doc.GetPDDoc()
                        file_path = pd_doc.GetFileName()
                        name = os.path.basename(file_path) if file_path else f"PDF_{i}"
                        results.append({
                            "app": "AcroExch.App",
                            "app_obj": acrobat,
                            "pd_doc": pd_doc,
                            "name": name,
                            "path": file_path or "",
                        })
                    except Exception:
                        continue
        except Exception:
            pass

        return results

    def detect_image_files(self, watch_dirs: list) -> list:
        """감시 폴더에서 이미지 파일 목록을 반환한다.

        Args:
            watch_dirs: 감시할 디렉토리 경로 리스트.

        Returns:
            이미지 파일 정보 dict 리스트.
            각 항목: {"app": "Image", "app_obj": 파일경로, "name": 파일명, "path": 전체경로}
        """
        IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif"}
        results = []
        for dir_path in watch_dirs:
            dir_path = os.path.abspath(dir_path)
            if not os.path.isdir(dir_path):
                continue
            for fname in os.listdir(dir_path):
                ext = os.path.splitext(fname)[1].lower()
                if ext in IMAGE_EXTS:
                    full_path = os.path.abspath(os.path.join(dir_path, fname))
                    results.append({
                        "app": "Image",
                        "app_obj": full_path,
                        "name": fname,
                        "path": full_path,
                    })
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
