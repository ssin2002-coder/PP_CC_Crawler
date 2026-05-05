"""
검증 API 라우트 모듈
검증 실행, 결과 조회, 결과 저장을 위한 REST API를 제공합니다.
"""

import logging
from flask import Blueprint, jsonify, request

import backend.storage as storage
from backend.excel_reader import get_excel_reader
from backend.rule_engine import RuleEngine
from backend.validators import VALIDATOR_MAP

logger = logging.getLogger(__name__)

validation_bp = Blueprint('validation', __name__)


def _get_rule_engine() -> RuleEngine:
    """RuleEngine 인스턴스를 반환합니다."""
    return RuleEngine(storage, VALIDATOR_MAP)


@validation_bp.route('/run', methods=['POST'])
def run_validation():
    """
    지정된 시트에 대해 검증을 실행합니다.

    Request body (JSON):
        {
            'workbook': str,
            'sheet': str,
            'max_row': int (optional, 기본 100),
            'max_col': int (optional, 기본 26),
            'rules': list (optional, 미제공 시 저장된 규칙 사용)
        }

    Response:
        200: {
            'issues': [...],
            'summary': {'errors', 'warnings', 'info', 'ok'},
            'total': int,
            'workbook': str,
            'sheet': str,
        }
    """
    body = request.get_json(silent=True) or {}
    workbook = body.get('workbook', '').strip()
    sheet = body.get('sheet', '').strip()
    max_row = int(body.get('max_row', 100))
    max_col = int(body.get('max_col', 26))
    custom_rules = body.get('rules')  # None이면 저장된 규칙 사용

    if not workbook or not sheet:
        return jsonify({'error': 'workbook, sheet 필드가 필요합니다.'}), 400

    try:
        # Excel 데이터 읽기
        reader = get_excel_reader()
        data = reader.read_range(workbook, sheet, max_row=max_row, max_col=max_col)
        cells = data.get('cells', [])

        if not cells:
            return jsonify({
                'issues': [],
                'summary': {'errors': 0, 'warnings': 0, 'info': 0, 'ok': True},
                'total': 0,
                'workbook': workbook,
                'sheet': sheet,
                'message': '데이터가 없습니다.',
            })

        # 이력 통계 컨텍스트 구성
        context = _build_context()

        # 검증 실행
        engine = _get_rule_engine()
        result = engine.run_validation(
            cells=cells,
            sheet=sheet,
            rules=custom_rules,
            context=context,
        )

        result['workbook'] = workbook
        result['sheet'] = sheet

        logger.info(
            f"검증 완료: {workbook}/{sheet} - "
            f"오류 {result['summary']['errors']}건, "
            f"경고 {result['summary']['warnings']}건"
        )

        return jsonify(result)

    except Exception as e:
        logger.error(f"검증 실행 오류 [{workbook}/{sheet}]: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@validation_bp.route('/results', methods=['GET'])
def list_results():
    """
    저장된 검증 결과 파일 목록을 반환합니다.

    Response:
        200: {
            'results': [
                {'filename': str, 'workbook': str, 'saved_at': str}
            ]
        }
    """
    try:
        filenames = storage.load_results()

        # 파일명에서 기본 메타데이터 추출
        results_meta = []
        for filename in filenames:
            result_data = storage.load_result(filename)
            if result_data:
                results_meta.append({
                    'filename': filename,
                    'workbook': result_data.get('workbook', ''),
                    'saved_at': result_data.get('saved_at', ''),
                    'summary': result_data.get('summary', {}),
                    'total': result_data.get('total', 0),
                })
            else:
                results_meta.append({'filename': filename, 'workbook': '', 'saved_at': ''})

        return jsonify({'results': results_meta, 'count': len(results_meta)})
    except Exception as e:
        logger.error(f"결과 목록 조회 오류: {e}")
        return jsonify({'error': str(e), 'results': []}), 500


@validation_bp.route('/export', methods=['POST'])
def export_result():
    """
    검증 결과를 파일로 저장합니다.

    Request body (JSON):
        {
            'workbook': str,
            'result': {issues, summary, total, ...}
        }

    Response:
        200: {'filename': str, 'message': str}
    """
    body = request.get_json(silent=True) or {}
    workbook = body.get('workbook', '').strip()
    result_data = body.get('result', {})

    if not workbook:
        return jsonify({'error': 'workbook 필드가 필요합니다.'}), 400

    if not result_data:
        return jsonify({'error': 'result 필드가 필요합니다.'}), 400

    try:
        filename = storage.save_result(workbook, result_data)
        return jsonify({
            'filename': filename,
            'message': f"검증 결과가 저장되었습니다: {filename}",
        })
    except Exception as e:
        logger.error(f"결과 저장 오류 [{workbook}]: {e}")
        return jsonify({'error': str(e)}), 500


@validation_bp.route('/results/<filename>', methods=['GET'])
def get_result(filename: str):
    """
    특정 검증 결과 파일을 반환합니다.

    Path params:
        filename: 결과 파일명

    Response:
        200: {result_data}
        404: {'error': '파일 없음'}
    """
    # 경로 탐색 방지
    if '/' in filename or '\\' in filename or '..' in filename:
        return jsonify({'error': '잘못된 파일명입니다.'}), 400

    try:
        result = storage.load_result(filename)
        if result is None:
            return jsonify({'error': f"결과 파일 없음: {filename}"}), 404
        return jsonify(result)
    except Exception as e:
        logger.error(f"결과 로드 오류 [{filename}]: {e}")
        return jsonify({'error': str(e)}), 500


def _build_context() -> dict:
    """
    검증에 사용할 컨텍스트를 구성합니다.
    이력 통계가 있으면 포함합니다.
    """
    context = {}
    try:
        from backend.history_manager import get_history_manager
        hm = get_history_manager()
        all_stats = hm.get_all_stats()
        if all_stats:
            context['history_stats'] = all_stats
    except Exception as e:
        logger.debug(f"이력 통계 로드 실패 (무시): {e}")
    return context
