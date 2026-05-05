"""
Excel COM 워커 (TCP 소켓 서버)
별도 프로세스로 실행되어 COM 작업을 수행합니다.
메인 프로세스와 localhost TCP 소켓으로 통신합니다.

Usage: python excel_com_worker.py <port>
"""

import socket
import json
import sys
import traceback
import datetime

XL_NONE = -4142


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 19876

    import pythoncom
    pythoncom.CoInitialize()

    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind(('127.0.0.1', port))
    server.listen(5)
    server.settimeout(1.0)

    print(f"[COM Worker] Listening on port {port}", flush=True)

    running = True
    while running:
        try:
            conn, addr = server.accept()
        except socket.timeout:
            continue
        except Exception:
            break

        try:
            # 요청 수신
            header = _recv_exact(conn, 8)
            if not header:
                conn.close()
                continue
            payload_len = int.from_bytes(header, 'big')
            payload = _recv_exact(conn, payload_len)
            if not payload:
                conn.close()
                continue

            request = json.loads(payload.decode('utf-8'))
            cmd = request.pop('cmd', '')

            if cmd == 'shutdown':
                _send_response(conn, {'data': 'ok'})
                conn.close()
                running = False
                continue

            # 명령 실행
            try:
                result = _execute(cmd, **request)
                _send_response(conn, {'data': result})
            except Exception as e:
                _send_response(conn, {'error': str(e)})

            conn.close()
        except Exception as e:
            try:
                conn.close()
            except Exception:
                pass

    server.close()
    pythoncom.CoUninitialize()
    print("[COM Worker] Shutdown", flush=True)


def _recv_exact(sock, n):
    data = b''
    while len(data) < n:
        chunk = sock.recv(n - len(data))
        if not chunk:
            return None
        data += chunk
    return data


def _send_response(conn, data):
    payload = json.dumps(data, default=_json_serial).encode('utf-8')
    header = len(payload).to_bytes(8, 'big')
    conn.sendall(header + payload)


def _json_serial(obj):
    """datetime 등 JSON 직렬화 불가 타입 처리"""
    if isinstance(obj, datetime.datetime):
        return obj.isoformat()
    return str(obj)


def _execute(command, **kwargs):
    handlers = {
        'get_open_workbooks': _get_open_workbooks,
        'get_sheets': _get_sheets,
        'read_range': _read_range,
        'get_status': _get_status,
        'navigate': _navigate,
    }
    handler = handlers.get(command)
    if not handler:
        raise ValueError(f"Unknown command: {command}")
    return handler(**kwargs)


def _get_excel():
    import win32com.client
    return win32com.client.GetActiveObject("Excel.Application")


def _get_open_workbooks(**kwargs):
    try:
        xl = _get_excel()
        workbooks = []
        for wb in xl.Workbooks:
            sheets = [ws.Name for ws in wb.Sheets]
            workbooks.append({'name': wb.Name, 'path': wb.FullName, 'sheets': sheets})
        return workbooks
    except Exception:
        return []


def _get_sheets(workbook_name='', **kwargs):
    try:
        xl = _get_excel()
        wb = xl.Workbooks(workbook_name)
        return [ws.Name for ws in wb.Sheets]
    except Exception:
        return []


def _read_range(workbook_name='', sheet_name='', max_row=100, max_col=26, **kwargs):
    try:
        xl = _get_excel()
        wb = xl.Workbooks(workbook_name)
        ws = wb.Sheets(sheet_name)

        used = ws.UsedRange
        actual_rows = min(used.Rows.Count, max_row)
        actual_cols = min(used.Columns.Count, max_col)

        # 벌크 읽기: Range.Value로 2D 튜플 한 번에 가져옴
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(actual_rows, actual_cols))
        values = rng.Value
        if values is None:
            values = ()
        if not isinstance(values, tuple):
            values = ((values,),)
        elif values and not isinstance(values[0], tuple):
            values = (values,)

        # 섹션/소계 행 감지 (볼드 판정용)
        section_kws = ('소계', '합계', '계', '자재비', '인건비', '경비', '총 합계',
                       '[자재비]', '[인건비]', '[경비]', '정비 비용')
        bold_rows = set()
        for ri, row in enumerate(values):
            if row and row[0] and any(kw in str(row[0]) for kw in section_kws):
                bold_rows.add(ri)

        cells_matrix = []
        for ri, row in enumerate(values):
            row_data = []
            for ci, val in enumerate(row):
                is_num = isinstance(val, (int, float)) and not isinstance(val, bool)
                if val is None:
                    disp = ''
                elif is_num:
                    disp = f'{val:,.0f}' if isinstance(val, float) and val == int(val) else str(val)
                else:
                    disp = str(val)
                row_data.append({
                    'value': disp, 'raw_value': val,
                    'row': ri + 1, 'col': ci + 1,
                    'is_bold': ri in bold_rows,
                    'is_number': is_num,
                    'bg_color': None, 'font_color': None,
                })
            cells_matrix.append(row_data)

        return {
            'workbook': workbook_name,
            'sheet': sheet_name,
            'cells': cells_matrix,
            'row_count': actual_rows,
            'col_count': actual_cols,
        }
    except Exception:
        return {'workbook': workbook_name, 'sheet': sheet_name,
                'cells': [], 'row_count': 0, 'col_count': 0}


def _read_cell(cell, row, col):
    try:
        raw_value = cell.Value
        is_number = isinstance(raw_value, (int, float)) and not isinstance(raw_value, bool)
        display_value = cell.Text if cell.Text is not None else ''
        if raw_value is None:
            display_value = ''

        bg_color = None
        try:
            bg_raw = cell.Interior.Color
            if bg_raw != XL_NONE and cell.Interior.ColorIndex != XL_NONE:
                b = (int(bg_raw) >> 16) & 0xFF
                g = (int(bg_raw) >> 8) & 0xFF
                r = int(bg_raw) & 0xFF
                bg_color = f"#{r:02X}{g:02X}{b:02X}"
        except Exception:
            pass

        return {
            'value': display_value,
            'raw_value': raw_value,
            'row': row, 'col': col,
            'is_bold': bool(cell.Font.Bold),
            'is_number': is_number,
            'bg_color': bg_color,
            'font_color': None,
        }
    except Exception:
        return {'value': '', 'raw_value': None, 'row': row, 'col': col,
                'is_bold': False, 'is_number': False, 'bg_color': None, 'font_color': None}


def _get_status(**kwargs):
    try:
        workbooks = _get_open_workbooks()
        return {'connected': len(workbooks) > 0, 'workbooks': workbooks}
    except Exception:
        return {'connected': False, 'workbooks': []}


def _navigate(workbook_name='', sheet_name='', cell_ref='', **kwargs):
    try:
        import win32com.client
        xl = win32com.client.GetActiveObject("Excel.Application")
        wb = xl.Workbooks(workbook_name)
        wb.Activate()
        ws = wb.Sheets(sheet_name)
        ws.Activate()
        ws.Range(cell_ref).Select()
        xl.Visible = True
        return {'success': True, 'message': f'{sheet_name}!{cell_ref}로 이동했습니다.'}
    except Exception as e:
        return {'success': False, 'message': str(e)}


if __name__ == '__main__':
    main()
