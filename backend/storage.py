"""
스토리지 모듈
규칙, 검증 결과를 JSON 파일로 저장하고 불러옵니다.
자동 저장 디렉토리 생성도 담당합니다.
"""

import json
import logging
import os
from datetime import datetime
from typing import Any, Dict, List, Optional

from backend.config import (
    AUTOSAVE_DIR,
    CUSTOM_RULES_FILENAME,
    DEFAULT_RULES_FILENAME,
    HISTORY_DIR,
    RESULTS_DIR,
    RULES_DIR,
)

logger = logging.getLogger(__name__)


def ensure_dirs() -> None:
    """
    필요한 모든 자동 저장 서브 디렉토리를 생성합니다.
    이미 존재하는 경우 무시합니다.
    """
    dirs = [AUTOSAVE_DIR, RULES_DIR, RESULTS_DIR, HISTORY_DIR]
    for d in dirs:
        os.makedirs(d, exist_ok=True)
    logger.debug("자동 저장 디렉토리 확인 완료")


# 앱 시작 시 디렉토리 생성
ensure_dirs()


# ------------------------------------------------------------------
# 규칙 관련
# ------------------------------------------------------------------

def load_rules() -> List[Dict]:
    """
    기본 규칙과 사용자 정의 규칙을 병합하여 반환합니다.

    우선순위: 사용자 정의 규칙 > 기본 규칙
    같은 id의 규칙은 사용자 정의 규칙으로 덮어씁니다.

    Returns:
        List[dict]: 병합된 규칙 목록
    """
    default_rules = _load_json_file(os.path.join(RULES_DIR, DEFAULT_RULES_FILENAME), default=[])
    custom_rules = _load_json_file(os.path.join(RULES_DIR, CUSTOM_RULES_FILENAME), default=[])

    # id를 키로 하는 딕셔너리로 변환 후 병합
    merged: Dict[str, Dict] = {}
    for rule in default_rules:
        merged[rule['id']] = rule
    for rule in custom_rules:
        merged[rule['id']] = rule  # 사용자 규칙이 기본 규칙을 덮어씀

    return list(merged.values())


def save_rules(rules: List[Dict]) -> None:
    """
    사용자 정의 규칙을 custom_rules.json에 저장합니다.

    Args:
        rules: 저장할 규칙 목록
    """
    path = os.path.join(RULES_DIR, CUSTOM_RULES_FILENAME)
    _save_json_file(path, rules)
    logger.info(f"사용자 정의 규칙 저장 완료: {len(rules)}건")


def save_default_rules(rules: List[Dict]) -> None:
    """
    기본 규칙을 default_rules.json에 저장합니다.

    Args:
        rules: 저장할 기본 규칙 목록
    """
    path = os.path.join(RULES_DIR, DEFAULT_RULES_FILENAME)
    _save_json_file(path, rules)
    logger.info(f"기본 규칙 저장 완료: {len(rules)}건")


# ------------------------------------------------------------------
# 검증 결과 관련
# ------------------------------------------------------------------

def load_results() -> List[str]:
    """
    저장된 검증 결과 파일명 목록을 반환합니다.
    최신 파일 순으로 정렬합니다.

    Returns:
        List[str]: 결과 파일명 목록 (경로 없이 파일명만)
    """
    try:
        files = [
            f for f in os.listdir(RESULTS_DIR)
            if f.endswith('.json')
        ]
        # 파일명 기준 내림차순 정렬 (날짜 포함이므로 최신 순)
        files.sort(reverse=True)
        return files
    except Exception as e:
        logger.error(f"결과 목록 로드 실패: {e}")
        return []


def save_result(workbook_name: str, result_data: Dict) -> str:
    """
    검증 결과를 results/{date}_{workbook}.json 형식으로 저장합니다.

    Args:
        workbook_name: 통합 문서 이름 (확장자 포함)
        result_data: 저장할 결과 데이터

    Returns:
        str: 저장된 파일명
    """
    # 파일명에 사용 불가한 문자 제거
    safe_name = _sanitize_filename(workbook_name)
    date_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"{date_str}_{safe_name}.json"
    path = os.path.join(RESULTS_DIR, filename)

    # 메타데이터 추가
    save_data = {
        'saved_at': datetime.now().isoformat(),
        'workbook': workbook_name,
        **result_data,
    }

    _save_json_file(path, save_data)
    logger.info(f"검증 결과 저장 완료: {filename}")
    return filename


def load_result(filename: str) -> Optional[Dict]:
    """
    지정된 파일명의 검증 결과를 불러옵니다.

    Args:
        filename: 결과 파일명 (경로 없이 파일명만)

    Returns:
        dict | None: 결과 데이터, 파일이 없으면 None
    """
    path = os.path.join(RESULTS_DIR, filename)
    if not os.path.isfile(path):
        logger.warning(f"결과 파일 없음: {filename}")
        return None
    return _load_json_file(path, default=None)


# ------------------------------------------------------------------
# 내부 유틸리티
# ------------------------------------------------------------------

def _load_json_file(path: str, default: Any = None) -> Any:
    """JSON 파일을 읽어 반환합니다. 파일이 없거나 오류 시 default를 반환합니다."""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        logger.debug(f"파일 없음 (기본값 사용): {path}")
        return default
    except json.JSONDecodeError as e:
        logger.error(f"JSON 파싱 오류 [{path}]: {e}")
        return default
    except Exception as e:
        logger.error(f"파일 읽기 오류 [{path}]: {e}")
        return default


def _save_json_file(path: str, data: Any) -> None:
    """데이터를 JSON 파일로 저장합니다."""
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"파일 저장 오류 [{path}]: {e}")
        raise


def _sanitize_filename(name: str) -> str:
    """파일명에 사용 불가한 문자를 제거합니다."""
    invalid_chars = r'\/:*?"<>|'
    result = name
    for ch in invalid_chars:
        result = result.replace(ch, '_')
    # 확장자 제거
    result = os.path.splitext(result)[0]
    # 최대 50자
    return result[:50]
