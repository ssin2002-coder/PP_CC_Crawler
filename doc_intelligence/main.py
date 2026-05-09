"""
main.py — Doc Intelligence 메인 앱
DocIntelligenceApp: COM 폴링 + 파이프라인 실행
main(): tkinter GUI 또는 터미널 fallback 실행
"""
import logging
import threading
import time

from doc_intelligence.engine import Engine
from doc_intelligence.com_worker import ComWorker
from doc_intelligence.parsers import ExcelParser, WordParser, PowerPointParser, PdfParser, ImageParser
from doc_intelligence.fingerprint import Fingerprinter
from doc_intelligence.extractor import EntityExtractor

logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────
# COM ProgID → 파서 키 매핑
# ──────────────────────────────────────────────

_APP_TO_PARSER = {
    "Excel.Application": "excel",
    "Word.Application": "word",
    "PowerPoint.Application": "ppt",
    "AcroExch.App": "pdf",
}


class DocIntelligenceApp:
    """COM 폴링 기반 문서 처리 앱."""

    def __init__(self):
        self.engine = Engine(db_path="templates.db")
        self.com_worker = ComWorker()
        self.parsers = {
            "excel": ExcelParser(),
            "word": WordParser(),
            "ppt": PowerPointParser(),
            "pdf": PdfParser(),
            "image": ImageParser(),
        }
        self.engine.register(Fingerprinter())
        self.engine.register(EntityExtractor())
        self._polling = False
        self._poll_thread: threading.Thread | None = None
        # 이미 처리한 문서 경로 추적 (중복 처리 방지)
        self._seen_paths: set = set()

    # ──────────────────────────────────────────────
    # 폴링
    # ──────────────────────────────────────────────

    def start_polling(self, interval: int = 3) -> None:
        """별도 스레드에서 com_worker.detect_open_documents 폴링.
        새 문서 감지 시 _process_document 호출.
        """
        self._polling = True
        self._poll_thread = threading.Thread(
            target=self._poll_loop,
            args=(interval,),
            daemon=True,
            name="DocIntelligencePollThread",
        )
        self._poll_thread.start()
        logger.info("COM 폴링 시작 (interval=%ds)", interval)

    def stop_polling(self) -> None:
        """폴링을 중지한다."""
        self._polling = False
        if self._poll_thread and self._poll_thread.is_alive():
            self._poll_thread.join(timeout=5)
        logger.info("COM 폴링 중지")

    def _poll_loop(self, interval: int) -> None:
        """폴링 루프 — _polling이 True인 동안 반복 실행한다."""
        while self._polling:
            try:
                docs = self.com_worker.detect_open_documents()
                for doc_info in docs:
                    path = doc_info.get("path", "")
                    if path and path not in self._seen_paths:
                        self._seen_paths.add(path)
                        try:
                            self._process_document(doc_info)
                        except Exception as exc:
                            logger.exception("문서 처리 중 예외 — %s: %s", path, exc)
            except Exception as exc:
                logger.exception("COM 감지 중 예외: %s", exc)
            time.sleep(interval)

    # ──────────────────────────────────────────────
    # 문서 처리
    # ──────────────────────────────────────────────

    def _process_document(self, doc_info: dict) -> dict:
        """
        parser 선택 → COM 파싱 → engine.process → 결과 반환.

        doc_info: {"app": prog_id, "name": 문서명, "path": 전체경로}
        반환: engine.process context dict
        """
        app = doc_info.get("app", "")
        parser_key = _APP_TO_PARSER.get(app)

        if parser_key is None:
            logger.warning("지원하지 않는 앱 ProgID: %s", app)
            return {}

        parser = self.parsers.get(parser_key)
        if parser is None:
            logger.warning("파서를 찾을 수 없음: %s", parser_key)
            return {}

        com_app = self.com_worker.get_active_app(app)
        parsed_doc = parser.parse_from_com(com_app)
        context = self.engine.process(parsed_doc)
        logger.info("문서 처리 완료 — %s / 엔티티 %d개",
                    doc_info.get("name", ""), len(context.get("entities", [])))
        return context


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    )

    app = DocIntelligenceApp()

    try:
        import tkinter as tk
        root = tk.Tk()
        root.title("Doc Intelligence v0.1")
        root.geometry("1280x720")
        root.configure(bg="#0f1117")
        tk.Label(
            root,
            text="Doc Intelligence v0.1",
            fg="#58a6ff",
            bg="#0f1117",
            font=("맑은 고딕", 16),
        ).pack(pady=20)
        app.start_polling()
        root.mainloop()
    except Exception:
        app.start_polling()
        input("Enter를 누르면 종료합니다...")
    finally:
        app.stop_polling()


if __name__ == "__main__":
    main()
