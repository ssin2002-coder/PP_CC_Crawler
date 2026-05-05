"""
Excel Validator - 진입점 모듈
Flask 서버를 포트 5000에서 시작하고 기본 브라우저를 엽니다.
PyInstaller frozen 경로를 처리합니다.
"""

import sys
import os
import signal
import threading
import multiprocessing
import webbrowser
import logging

# PyInstaller frozen 실행 파일 경로 처리
if getattr(sys, 'frozen', False):
    # PyInstaller 번들 실행 시 임시 디렉토리를 BASE_DIR로 설정
    BASE_DIR = sys._MEIPASS
    # 실제 실행 파일이 있는 디렉토리를 작업 디렉토리로 설정
    os.chdir(os.path.dirname(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 프로젝트 루트를 sys.path에 추가
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)
logger = logging.getLogger(__name__)

SERVER_URL = "http://localhost:5000"


def open_browser() -> None:
    """서버 시작 후 기본 브라우저를 엽니다."""
    import time
    # 서버 초기화 대기
    time.sleep(1.5)
    webbrowser.open(SERVER_URL)
    logger.info(f"브라우저 열기: {SERVER_URL}")


def handle_shutdown(signum: int, frame) -> None:
    """Ctrl+C 등 종료 시그널을 처리합니다."""
    logger.info("서버 종료 중...")
    sys.exit(0)


def main() -> None:
    """Flask 서버를 시작하는 메인 함수입니다."""
    from backend.app import create_app

    # 종료 시그널 핸들러 등록
    signal.signal(signal.SIGINT, handle_shutdown)
    signal.signal(signal.SIGTERM, handle_shutdown)

    app, socketio = create_app()

    # 브라우저 열기는 별도 스레드에서 수행
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()

    logger.info(f"Excel Validator 서버 시작: {SERVER_URL}")

    try:
        # threading 모드로 실행 (COM 스레드 호환)
        socketio.run(
            app,
            host="0.0.0.0",
            port=5000,
            debug=False,
            use_reloader=False,
            log_output=False,
        )
    except KeyboardInterrupt:
        logger.info("키보드 인터럽트로 서버 종료")
    except Exception as e:
        logger.error(f"서버 오류: {e}")
        sys.exit(1)


if __name__ == "__main__":
    multiprocessing.freeze_support()  # PyInstaller 호환
    main()
