"""
Excel API 라우트 모듈
열려 있는 Excel 통합 문서의 데이터를 읽고 셀 탐색을 수행하는 REST API를 제공합니다.
"""

import logging
from flask import Blueprint, jsonify, request

from backend.excel_reader import get_excel_reader

from backend.excel_navigator import navigate_to_cell

logger = logging.getLogger(__name__)

excel_bp = Blueprint('excel', __name__)


@excel_bp.route('/workbooks', methods=['GET'])
def get_workbooks():
    """
    현재 열려 있는 Excel 통합 문서 목록을 반환합니다.

    Response:
        200: {'workbooks': [{'name', 'path', 'sheets'}, ...]}
    """
    try:
        reader = get_excel_reader()
        workbooks = reader.get_open_workbooks()
        return jsonify({'workbooks': workbooks})
    except Exception as e:
        logger.error(f"통합 문서 목록 조회 오류: {e}")
        return jsonify({'error': str(e), 'workbooks': []}), 500


@excel_bp.route('/sheets', methods=['GET'])
def get_sheets():
    """
    지정된 통합 문서의 시트 목록을 반환합니다.

    Query params:
        workbook (str): 통합 문서 이름

    Response:
        200: {'sheets': ['Sheet1', 'Sheet2', ...]}
        400: {'error': 'workbook 파라미터 필요'}
    """
    workbook = request.args.get('workbook', '').strip()
    if not workbook:
        return jsonify({'error': 'workbook 파라미터가 필요합니다.'}), 400

    try:
        reader = get_excel_reader()
        sheets = reader.get_sheets(workbook)
        return jsonify({'sheets': sheets})
    except Exception as e:
        logger.error(f"시트 목록 조회 오류 [{workbook}]: {e}")
        return jsonify({'error': str(e), 'sheets': []}), 500


@excel_bp.route('/data', methods=['GET'])
def get_data():
    """
    지정된 시트의 셀 데이터 매트릭스를 반환합니다.

    Query params:
        workbook (str): 통합 문서 이름
        sheet (str): 시트 이름
        max_row (int, optional): 최대 행 수 (기본 100)
        max_col (int, optional): 최대 열 수 (기본 26)

    Response:
        200: {
            'workbook': str,
            'sheet': str,
            'cells': [[cell, ...], ...],
            'row_count': int,
            'col_count': int,
            'formatting': {'theme': 'dark'}
        }
    """
    workbook = request.args.get('workbook', '').strip()
    sheet = request.args.get('sheet', '').strip()
    max_row = int(request.args.get('max_row', 100))
    max_col = int(request.args.get('max_col', 26))

    if not workbook:
        return jsonify({'error': 'workbook 파라미터가 필요합니다.'}), 400
    if not sheet:
        return jsonify({'error': 'sheet 파라미터가 필요합니다.'}), 400

    try:
        reader = get_excel_reader()
        data = reader.read_range(workbook, sheet, max_row=max_row, max_col=max_col)
        # 다크 테마 서식 메타데이터 추가
        data['formatting'] = {
            'theme': 'dark',
            'highlight_error': '#FF4444',
            'highlight_warning': '#FFA500',
            'highlight_info': '#4488FF',
        }
        return jsonify(data)
    except Exception as e:
        logger.error(f"셀 데이터 조회 오류 [{workbook}/{sheet}]: {e}")
        return jsonify({'error': str(e)}), 500


@excel_bp.route('/navigate', methods=['POST'])
def navigate_cell():
    """
    Excel에서 지정된 셀로 이동합니다.

    Request body (JSON):
        {
            'workbook': str,
            'sheet': str,
            'cell': str  (예: 'A1')
        }

    Response:
        200: {'success': bool, 'message': str}
    """
    body = request.get_json(silent=True) or {}
    workbook = body.get('workbook', '').strip()
    sheet = body.get('sheet', '').strip()
    cell = body.get('cell', '').strip()

    if not workbook or not sheet or not cell:
        return jsonify({'error': 'workbook, sheet, cell 필드가 모두 필요합니다.'}), 400

    try:
        result = navigate_to_cell(workbook, sheet, cell)
        return jsonify(result)
    except Exception as e:
        logger.error(f"셀 탐색 오류 [{workbook}/{sheet}/{cell}]: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@excel_bp.route('/status', methods=['GET'])
def get_status():
    """
    Excel 연결 상태를 반환합니다.

    Response:
        200: {
            'connected': bool,
            'workbooks': [...],
            'version': str
        }
    """
    try:
        reader = get_excel_reader()
        status = reader.get_status()
        status['version'] = '1.0.0'
        return jsonify(status)
    except Exception as e:
        logger.error(f"상태 조회 오류: {e}")
        return jsonify({'connected': False, 'workbooks': [], 'error': str(e)}), 500
