"""
규칙 관리 API 라우트 모듈
검증 규칙의 CRUD 및 토글을 지원하는 REST API를 제공합니다.
"""

import logging
import uuid
from flask import Blueprint, jsonify, request

import backend.storage as storage

logger = logging.getLogger(__name__)

rule_bp = Blueprint('rules', __name__)


@rule_bp.route('', methods=['GET'])
def get_rules():
    """
    모든 규칙(기본 + 사용자 정의 병합)을 반환합니다.

    Response:
        200: {'rules': [...], 'count': int}
    """
    try:
        rules = storage.load_rules()
        return jsonify({'rules': rules, 'count': len(rules)})
    except Exception as e:
        logger.error(f"규칙 목록 조회 오류: {e}")
        return jsonify({'error': str(e), 'rules': []}), 500


@rule_bp.route('', methods=['POST'])
def create_rule():
    """
    새 규칙을 생성하고 저장합니다.

    Request body (JSON): 규칙 딕셔너리 (id 없어도 자동 생성)

    Response:
        201: {'rule': {...}}
        400: {'error': '필수 필드 없음'}
    """
    body = request.get_json(silent=True) or {}

    # 필수 필드 검증
    required_fields = ('name', 'template')
    missing = [f for f in required_fields if not body.get(f)]
    if missing:
        return jsonify({'error': f"필수 필드 누락: {', '.join(missing)}"}), 400

    # ID 자동 생성
    if not body.get('id'):
        body['id'] = f"custom_{uuid.uuid4().hex[:8]}"

    # 기본값 설정
    body.setdefault('type', 'custom')
    body.setdefault('enabled', True)
    body.setdefault('severity', 'warning')
    body.setdefault('config', {})

    try:
        # 기존 사용자 규칙 로드 후 추가 (custom_rules.json만)
        custom_rules = _load_custom_rules()
        custom_rules.append(body)
        storage.save_rules(custom_rules)

        logger.info(f"규칙 생성 완료: {body['id']}")
        return jsonify({'rule': body}), 201
    except Exception as e:
        logger.error(f"규칙 생성 오류: {e}")
        return jsonify({'error': str(e)}), 500


@rule_bp.route('/<rule_id>', methods=['PUT'])
def update_rule(rule_id: str):
    """
    기존 규칙을 수정합니다.

    Path params:
        rule_id: 수정할 규칙의 ID

    Request body (JSON): 수정할 필드

    Response:
        200: {'rule': {...}}
        404: {'error': '규칙 없음'}
    """
    body = request.get_json(silent=True) or {}
    body['id'] = rule_id  # URL의 id를 우선 사용

    try:
        custom_rules = _load_custom_rules()
        rule_idx = _find_rule_index(custom_rules, rule_id)

        if rule_idx is not None:
            # 기존 필드 유지하면서 업데이트
            custom_rules[rule_idx].update(body)
            updated_rule = custom_rules[rule_idx]
        else:
            # custom_rules에 없으면 새로 추가 (기본 규칙 오버라이드)
            custom_rules.append(body)
            updated_rule = body

        storage.save_rules(custom_rules)
        logger.info(f"규칙 수정 완료: {rule_id}")
        return jsonify({'rule': updated_rule})
    except Exception as e:
        logger.error(f"규칙 수정 오류 [{rule_id}]: {e}")
        return jsonify({'error': str(e)}), 500


@rule_bp.route('/<rule_id>', methods=['DELETE'])
def delete_rule(rule_id: str):
    """
    규칙을 삭제합니다.

    Path params:
        rule_id: 삭제할 규칙의 ID

    Response:
        200: {'deleted': str}
        404: {'error': '규칙 없음'}
    """
    try:
        custom_rules = _load_custom_rules()
        rule_idx = _find_rule_index(custom_rules, rule_id)

        if rule_idx is None:
            return jsonify({'error': f"규칙을 찾을 수 없음: {rule_id}"}), 404

        custom_rules.pop(rule_idx)
        storage.save_rules(custom_rules)

        logger.info(f"규칙 삭제 완료: {rule_id}")
        return jsonify({'deleted': rule_id})
    except Exception as e:
        logger.error(f"규칙 삭제 오류 [{rule_id}]: {e}")
        return jsonify({'error': str(e)}), 500


@rule_bp.route('/<rule_id>/toggle', methods=['PATCH'])
def toggle_rule(rule_id: str):
    """
    규칙의 활성화 상태를 토글합니다.

    Path params:
        rule_id: 토글할 규칙의 ID

    Response:
        200: {'rule_id': str, 'enabled': bool}
    """
    try:
        custom_rules = _load_custom_rules()
        rule_idx = _find_rule_index(custom_rules, rule_id)

        if rule_idx is not None:
            # custom_rules에 있는 경우 토글
            current = custom_rules[rule_idx].get('enabled', True)
            custom_rules[rule_idx]['enabled'] = not current
            new_state = custom_rules[rule_idx]['enabled']
        else:
            # custom_rules에 없으면 기본 규칙을 오버라이드로 추가
            # 기본 규칙에서 현재 상태 확인
            all_rules = storage.load_rules()
            source_rule = next((r for r in all_rules if r['id'] == rule_id), None)

            if source_rule is None:
                return jsonify({'error': f"규칙을 찾을 수 없음: {rule_id}"}), 404

            current = source_rule.get('enabled', True)
            new_state = not current
            custom_rules.append({**source_rule, 'enabled': new_state})

        storage.save_rules(custom_rules)
        logger.info(f"규칙 토글 완료: {rule_id} -> enabled={new_state}")
        return jsonify({'rule_id': rule_id, 'enabled': new_state})
    except Exception as e:
        logger.error(f"규칙 토글 오류 [{rule_id}]: {e}")
        return jsonify({'error': str(e)}), 500


# ------------------------------------------------------------------
# 내부 유틸리티
# ------------------------------------------------------------------

def _load_custom_rules():
    """custom_rules.json만 로드합니다 (병합 없이)."""
    import os
    import json
    from backend.config import RULES_DIR, CUSTOM_RULES_FILENAME

    path = os.path.join(RULES_DIR, CUSTOM_RULES_FILENAME)
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []
    except Exception as e:
        logger.error(f"custom_rules.json 로드 오류: {e}")
        return []


def _find_rule_index(rules: list, rule_id: str):
    """규칙 목록에서 ID로 인덱스를 찾습니다."""
    for idx, rule in enumerate(rules):
        if rule.get('id') == rule_id:
            return idx
    return None
