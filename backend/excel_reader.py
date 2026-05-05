"""
Excel COM 리더 모듈 (소켓 IPC 기반)
COM 워커를 별도 프로세스로 실행하고 localhost TCP 소켓으로 통신합니다.
Flask/eventlet/threading 간섭을 완전 차단합니다.
"""

import subprocess
import socket
import json
import logging
import threading
import sys
import os
from typing import Any, Callable, Dict, List, Optional

logger = logging.getLogger(__name__)

COM_WORKER_PORT = 19876


class ExcelReader:
    """소켓 IPC로 COM 워커 프로세스와 통신"""

    def __init__(self) -> None:
        self._worker_proc = None
        self._running = False
        self._dispatch_lock = threading.Lock()
        self._change_callback: Optional[Callable] = None
        self._prev_snapshot: str = ""

    def start(self) -> None:
        if self._running:
            return
        worker_script = os.path.join(os.path.dirname(__file__), 'excel_com_worker.py')
        python_exe = sys.executable
        self._worker_proc = subprocess.Popen(
            [python_exe, worker_script, str(COM_WORKER_PORT)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        self._running = True
        # 워커 시작 대기
        import time
        for _ in range(30):
            try:
                s = socket.create_connection(('127.0.0.1', COM_WORKER_PORT), timeout=1)
                s.close()
                break
            except Exception:
                time.sleep(0.2)
        logger.info(f"ExcelReader COM 워커 시작 (PID: {self._worker_proc.pid}, PORT: {COM_WORKER_PORT})")

    def stop(self) -> None:
        self._running = False
        try:
            self._send_command('shutdown')
        except Exception:
            pass
        if self._worker_proc:
            self._worker_proc.terminate()

    def set_change_callback(self, callback: Callable) -> None:
        self._change_callback = callback

    def get_open_workbooks(self) -> List[Dict]:
        return self._send_command('get_open_workbooks') or []

    def get_sheets(self, workbook_name: str) -> List[str]:
        return self._send_command('get_sheets', workbook_name=workbook_name) or []

    def read_range(self, workbook_name: str, sheet_name: str,
                   max_row: int = 100, max_col: int = 26) -> Dict:
        result = self._send_command('read_range',
                                    workbook_name=workbook_name,
                                    sheet_name=sheet_name,
                                    max_row=max_row, max_col=max_col)
        return result or {'workbook': workbook_name, 'sheet': sheet_name,
                          'cells': [], 'row_count': 0, 'col_count': 0}

    def get_status(self) -> Dict:
        result = self._send_command('get_status')
        return result or {'connected': False, 'workbooks': []}

    def navigate_to_cell(self, workbook_name: str, sheet_name: str, cell_ref: str) -> Dict:
        result = self._send_command('navigate',
                                    workbook_name=workbook_name,
                                    sheet_name=sheet_name,
                                    cell_ref=cell_ref)
        return result or {'success': False, 'message': 'timeout'}

    def poll_changes(self) -> Optional[Dict]:
        try:
            workbooks = self.get_open_workbooks()
            if not workbooks:
                if self._prev_snapshot:
                    self._prev_snapshot = ""
                    if self._change_callback:
                        self._change_callback({'type': 'disconnected'})
                return None

            wb = workbooks[0]
            sheet = wb['sheets'][0] if wb.get('sheets') else None
            if not sheet:
                return None

            data = self.read_range(wb['name'], sheet)
            snapshot = str([[c.get('value') for c in row] for row in data.get('cells', [])])

            if snapshot != self._prev_snapshot:
                self._prev_snapshot = snapshot
                if self._change_callback:
                    self._change_callback({'type': 'data_changed', 'data': data})
                return data
            return None
        except Exception as e:
            logger.debug(f"변경 감지 오류: {e}")
            return None

    def _send_command(self, command: str, **kwargs) -> Any:
        """소켓으로 COM 워커에 명령 전송 후 결과 수신"""
        with self._dispatch_lock:
            try:
                s = socket.create_connection(('127.0.0.1', COM_WORKER_PORT), timeout=10)
                payload = json.dumps({'cmd': command, **kwargs}).encode('utf-8')
                # 길이 헤더 (8바이트) + 페이로드
                header = len(payload).to_bytes(8, 'big')
                s.sendall(header + payload)

                # 응답 수신
                resp_header = self._recv_exact(s, 8)
                if not resp_header:
                    return None
                resp_len = int.from_bytes(resp_header, 'big')
                resp_data = self._recv_exact(s, resp_len)
                s.close()

                if resp_data:
                    result = json.loads(resp_data.decode('utf-8'))
                    if result.get('error'):
                        logger.debug(f"COM 오류: {result['error']}")
                        return None
                    return result.get('data')
                return None
            except socket.timeout:
                logger.warning(f"COM 타임아웃: {command}")
                return None
            except Exception as e:
                logger.debug(f"COM 통신 오류: {e}")
                return None

    def _recv_exact(self, sock, n):
        data = b''
        while len(data) < n:
            chunk = sock.recv(n - len(data))
            if not chunk:
                return None
            data += chunk
        return data


_excel_reader: Optional[ExcelReader] = None

def get_excel_reader() -> ExcelReader:
    global _excel_reader
    if _excel_reader is None:
        _excel_reader = ExcelReader()
        _excel_reader.start()
    return _excel_reader
