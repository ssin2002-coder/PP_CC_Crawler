"""
Flask 앱 팩토리 모듈
CORS, SocketIO, 블루프린트를 설정하고 React SPA를 서빙합니다.
"""

import os
import logging
from typing import Tuple

from flask import Flask, send_from_directory, jsonify
from flask_cors import CORS
from flask_socketio import SocketIO

from backend.config import FRONTEND_DIST, PORT, POLL_INTERVAL

logger = logging.getLogger(__name__)

# SocketIO 인스턴스 (전역 - ws_events에서 import하여 사용)
socketio = SocketIO()


def create_app() -> Tuple[Flask, SocketIO]:
    """
    Flask 앱을 생성하고 구성합니다.

    Returns:
        Tuple[Flask, SocketIO]: Flask 앱과 SocketIO 인스턴스
    """
    app = Flask(__name__, static_folder=None)
    app.config['SECRET_KEY'] = 'excel-validator-secret-key-2025'
    app.config['DEBUG'] = False

    # CORS 설정 - 개발 환경에서 React dev server(3000 포트)도 허용
    CORS(app, resources={r"/api/*": {"origins": "*"}})

    # SocketIO 초기화 (threading 모드 — COM 스레드와의 호환성)
    socketio.init_app(
        app,
        cors_allowed_origins="*",
        async_mode='threading',
        logger=False,
        engineio_logger=False,
    )

    # 블루프린트 등록
    _register_blueprints(app)

    # SocketIO 이벤트 핸들러 등록
    _register_ws_events(app)

    # React SPA 서빙 라우트 등록
    _register_frontend_routes(app)

    # 앱 시작 시 Excel 폴링 스레드 시작
    _start_excel_polling()

    return app, socketio


def _register_blueprints(app: Flask) -> None:
    """API 블루프린트를 등록합니다."""
    from backend.api.excel_routes import excel_bp
    from backend.api.rule_routes import rule_bp
    from backend.api.validation_routes import validation_bp

    app.register_blueprint(excel_bp, url_prefix='/api/excel')
    app.register_blueprint(rule_bp, url_prefix='/api/rules')
    app.register_blueprint(validation_bp, url_prefix='/api/validate')

    logger.info("API 블루프린트 등록 완료")


def _register_ws_events(app: Flask) -> None:
    """SocketIO 이벤트 핸들러를 등록합니다."""
    # ws_events 모듈 임포트로 핸들러가 자동 등록됨
    from backend.api import ws_events  # noqa: F401
    logger.info("WebSocket 이벤트 핸들러 등록 완료")


def _register_frontend_routes(app: Flask) -> None:
    """React SPA 정적 파일 서빙 라우트를 등록합니다."""

    @app.route('/', defaults={'path': ''})
    @app.route('/<path:path>')
    def serve_frontend(path: str):
        """
        React SPA를 서빙합니다.
        정적 파일이 존재하면 해당 파일을, 아니면 index.html을 반환합니다.
        """
        # frontend/dist 디렉토리가 없으면 개발 모드 안내
        if not os.path.isdir(FRONTEND_DIST):
            return jsonify({
                'status': 'backend_only',
                'message': 'Frontend not built. Run: cd frontend && npm run build',
                'api_docs': '/api/excel/status',
            }), 200

        # 정적 파일 요청인 경우 해당 파일 반환
        target_file = os.path.join(FRONTEND_DIST, path)
        if path and os.path.isfile(target_file):
            return send_from_directory(FRONTEND_DIST, path)

        # SPA 라우팅: 나머지 모든 요청에 index.html 반환
        index_path = os.path.join(FRONTEND_DIST, 'index.html')
        if os.path.isfile(index_path):
            return send_from_directory(FRONTEND_DIST, 'index.html')

        return jsonify({'error': 'index.html not found in frontend/dist'}), 404


def _start_excel_polling() -> None:
    """앱 시작 시 Excel 폴링 백그라운드 스레드를 시작합니다."""
    from backend.api.ws_events import start_polling_thread
    start_polling_thread(socketio, POLL_INTERVAL)
    logger.info(f"Excel 폴링 스레드 시작 (간격: {POLL_INTERVAL}초)")
