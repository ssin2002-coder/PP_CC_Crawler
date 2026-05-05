"""
WebSocket 이벤트 핸들러 모듈
SocketIO 연결, Excel 실시간 데이터, 검증 실행 이벤트를 처리합니다.
백그라운드 스레드가 4초마다 Excel 변경을 감지하고 클라이언트에 푸시합니다.
"""

import logging
import threading
import time
from typing import Optional

from flask_socketio import SocketIO, emit

logger = logging.getLogger(__name__)

# 폴링 스레드 전역 참조
_polling_thread: Optional[threading.Thread] = None
_polling_active = False


def register_events(socketio: SocketIO) -> None:
    """
    SocketIO 이벤트 핸들러를 등록합니다.

    Args:
        socketio: Flask-SocketIO 인스턴스
    """

    @socketio.on('connect')
    def handle_connect():
        """클라이언트 연결 시 현재 Excel 상태를 전송합니다."""
        logger.info("WebSocket 클라이언트 연결")
        try:
            from backend.excel_reader import get_excel_reader
            reader = get_excel_reader()
            status = reader.get_status()
            emit('excel:status', status)
        except Exception as e:
            logger.error(f"연결 시 상태 전송 오류: {e}")
            emit('excel:status', {'connected': False, 'workbooks': []})

    @socketio.on('disconnect')
    def handle_disconnect():
        """클라이언트 연결 해제 시 로그를 기록합니다."""
        logger.info("WebSocket 클라이언트 연결 해제")

    @socketio.on('excel:refresh')
    def handle_refresh(data=None):
        """
        클라이언트의 데이터 갱신 요청을 처리합니다.
        workbook, sheet를 받아 최신 셀 데이터를 반환합니다.

        Event data (optional):
            {workbook: str, sheet: str}
        """
        data = data or {}
        workbook = data.get('workbook', '').strip()
        sheet = data.get('sheet', '').strip()

        try:
            from backend.excel_reader import get_excel_reader
            reader = get_excel_reader()

            if workbook and sheet:
                # 특정 시트 데이터 반환
                cell_data = reader.read_range(workbook, sheet)
                emit('excel:data', cell_data)
            else:
                # 전체 상태 반환
                status = reader.get_status()
                emit('excel:status', status)

        except Exception as e:
            logger.error(f"excel:refresh 처리 오류: {e}")
            emit('excel:error', {'message': str(e)})

    @socketio.on('validation:run')
    def handle_validation_run(data=None):
        """
        클라이언트의 검증 실행 요청을 처리합니다.

        Event data:
            {workbook: str, sheet: str, rules: list (optional)}
        """
        data = data or {}
        workbook = data.get('workbook', '').strip()
        sheet = data.get('sheet', '').strip()
        custom_rules = data.get('rules')

        if not workbook or not sheet:
            emit('validation:error', {'message': 'workbook과 sheet가 필요합니다.'})
            return

        try:
            from backend.excel_reader import get_excel_reader
            from backend.rule_engine import RuleEngine
            from backend.validators import VALIDATOR_MAP
            import backend.storage as storage

            reader = get_excel_reader()
            cell_data = reader.read_range(workbook, sheet)
            cells = cell_data.get('cells', [])

            engine = RuleEngine(storage, VALIDATOR_MAP)
            result = engine.run_validation(
                cells=cells,
                sheet=sheet,
                rules=custom_rules,
            )
            result['workbook'] = workbook
            result['sheet'] = sheet

            emit('validation:result', result)
            logger.info(f"WebSocket 검증 완료: {workbook}/{sheet}")

        except Exception as e:
            logger.error(f"validation:run 처리 오류: {e}", exc_info=True)
            emit('validation:error', {'message': str(e)})


def start_polling_thread(socketio: SocketIO, interval: int = 4) -> None:
    """
    Excel 변경 감지 폴링 스레드를 시작합니다.

    Args:
        socketio: Flask-SocketIO 인스턴스
        interval: 폴링 간격 (초)
    """
    global _polling_thread, _polling_active

    # 이미 실행 중이면 중복 시작하지 않음
    if _polling_active and _polling_thread and _polling_thread.is_alive():
        logger.debug("폴링 스레드 이미 실행 중")
        return

    # 이벤트 핸들러 등록
    register_events(socketio)

    _polling_active = True
    _polling_thread = threading.Thread(
        target=_polling_loop,
        args=(socketio, interval),
        name="ExcelPollingThread",
        daemon=True,
    )
    _polling_thread.start()
    logger.info(f"Excel 폴링 스레드 시작 (간격: {interval}초)")


def _run_auto_validation(socketio: SocketIO, workbook_name: str, sheet_name: str, cells: list) -> None:
    """데이터 감지 후 자동 검증을 실행하고 결과를 클라이언트에 전송합니다."""
    try:
        from backend.rule_engine import RuleEngine
        from backend.validators import VALIDATOR_MAP
        import backend.storage as storage

        engine = RuleEngine(storage, VALIDATOR_MAP)
        result = engine.run_validation(cells=cells, sheet=sheet_name)
        result['workbook'] = workbook_name
        result['sheet'] = sheet_name

        socketio.emit('validation:result', result)
        logger.info(
            f"자동 검증 완료: {workbook_name}/{sheet_name} - "
            f"오류 {result['summary']['errors']}건, 경고 {result['summary']['warnings']}건"
        )
    except Exception as e:
        logger.error(f"자동 검증 오류: {e}")


def _polling_loop(socketio: SocketIO, interval: int) -> None:
    """
    주기적으로 Excel 상태와 데이터 변경을 확인하고 클라이언트에 푸시하는 루프입니다.
    새 워크북 감지 시: excel:status + 셀 데이터 + 자동 검증
    데이터 변경 시: excel:data_changed + 자동 검증
    """
    global _polling_active

    time.sleep(2)
    logger.info("Excel 폴링 루프 시작")

    prev_workbook_names: set = set()
    callback_set = False

    from backend.excel_reader import get_excel_reader

    while _polling_active:
        try:
            reader = get_excel_reader()

            # 콜백은 한 번만 설정
            if not callback_set:
                def on_change(event: dict) -> None:
                    event_type = event.get('type', 'unknown')
                    if event_type == 'data_changed':
                        change_data = event.get('data', {})
                        socketio.emit('excel:data_changed', change_data)
                    elif event_type == 'disconnected':
                        socketio.emit('excel:status', {'connected': False, 'workbooks': []})

                reader.set_change_callback(on_change)
                callback_set = True

            status = reader.get_status()
            if status is None:
                time.sleep(interval)
                continue

            current_names = {wb['name'] for wb in status.get('workbooks', [])}

            # 워크북 목록 변경 감지
            if current_names != prev_workbook_names:
                logger.info(f"워크북 변경 감지: {prev_workbook_names} -> {current_names}")
                socketio.emit('excel:status', status)

                # 새로 열린 워크북 → 데이터 + 자동 검증
                new_names = current_names - prev_workbook_names
                for wb in status.get('workbooks', []):
                    if wb['name'] in new_names and wb.get('sheets'):
                        sheet = wb['sheets'][0]
                        try:
                            data = reader.read_range(wb['name'], sheet)
                            cells = data.get('cells', [])
                            socketio.emit('excel:data_changed', data)
                            if cells:
                                _run_auto_validation(socketio, wb['name'], sheet, cells)
                        except Exception as e:
                            logger.error(f"신규 워크북 처리 오류: {e}")

                prev_workbook_names = current_names

            # 데이터 변경 감지
            reader.poll_changes()

        except Exception as e:
            logger.debug(f"폴링 오류 (무시): {e}")

        time.sleep(interval)

    logger.info("Excel 폴링 루프 종료")


def stop_polling() -> None:
    """폴링 스레드를 중지합니다."""
    global _polling_active
    _polling_active = False
    logger.info("Excel 폴링 중지 요청")
