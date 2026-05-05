"""
범위 검증기
지정된 열의 값이 허용 범위 [min, max] 안에 있는지 검증합니다.
"""

import logging
import uuid
from typing import Dict, List, Optional, Tuple

from backend.validators.base import BaseValidator, Issue

logger = logging.getLogger(__name__)

# 기본 열 범위 설정
DEFAULT_COLUMN_RANGES: Dict[str, Tuple[float, float]] = {
    '시간': (0, 24),
    '공수': (0, 24),
    '수량': (0, 999999),
}


class RangeValidator(BaseValidator):
    """
    숫자 값이 지정된 범위 내에 있는지 확인하는 검증기.

    헤더 이름 또는 열 인덱스로 검사 열을 지정합니다.
    """

    def validate(
        self,
        cells: list,
        rule_config: dict,
        context: dict = None,
    ) -> List[Issue]:
        """
        범위 검증을 수행합니다.

        rule_config 주요 키:
            column_ranges (dict): {열 헤더 또는 인덱스: {min: float, max: float}}
                예: {"시간": {"min": 0, "max": 24}}
            use_defaults (bool): 기본 범위 설정 사용 여부 (기본 True)

        Returns:
            List[Issue]: 범위 초과 이슈 목록
        """
        issues: List[Issue] = []
        column_ranges = rule_config.get('column_ranges', {})
        use_defaults = rule_config.get('use_defaults', True)
        rule_id = rule_config.get('id', 'auto_range')
        rule_name = rule_config.get('name', '범위 검증')
        severity = rule_config.get('severity', 'warning')
        sheet = rule_config.get('sheet', '')

        if not cells or len(cells) < 2:
            return issues

        # 헤더 행 파싱
        header_map = self._parse_header(cells[0])

        # 기본 범위 + 사용자 범위 병합
        effective_ranges: Dict[str, Tuple[Optional[float], Optional[float]]] = {}

        if use_defaults:
            for header_kw, (min_v, max_v) in DEFAULT_COLUMN_RANGES.items():
                # 헤더에서 키워드 매칭
                for header_text, col_idx in header_map.items():
                    if header_kw in header_text:
                        key = str(col_idx)
                        effective_ranges[key] = (min_v, max_v)

        # 사용자 정의 범위 (열 이름 또는 인덱스로 지정)
        for col_key, range_cfg in column_ranges.items():
            min_v = range_cfg.get('min')
            max_v = range_cfg.get('max')
            # 열 이름이면 인덱스로 변환
            if not str(col_key).isdigit():
                col_idx = header_map.get(str(col_key))
                if col_idx is not None:
                    effective_ranges[str(col_idx)] = (min_v, max_v)
            else:
                effective_ranges[str(col_key)] = (min_v, max_v)

        # 데이터 행 검사
        for row_idx, row in enumerate(cells[1:], start=1):
            for col_key, (min_v, max_v) in effective_ranges.items():
                col_idx = int(col_key)
                if col_idx >= len(row):
                    continue

                cell = row[col_idx]
                val = self._get_numeric_value(cell)
                if val is None:
                    continue

                out_of_range = False
                range_desc = ''

                if min_v is not None and val < min_v:
                    out_of_range = True
                    range_desc = f"최솟값 {min_v} 미만"
                elif max_v is not None and val > max_v:
                    out_of_range = True
                    range_desc = f"최댓값 {max_v} 초과"

                if out_of_range:
                    cell_ref = self._cell_ref(row_idx + 1, col_idx + 1)
                    expected = f"[{min_v}, {max_v}]"
                    issues.append(Issue(
                        id=str(uuid.uuid4()),
                        rule_id=rule_id,
                        rule_name=rule_name,
                        severity=severity,
                        cell_ref=cell_ref,
                        sheet=sheet,
                        message=f"범위 초과: {val} ({range_desc}, 허용 범위 {expected})",
                        current_value=str(val),
                        expected_value=expected,
                    ))

        logger.debug(f"범위 검증 완료: {len(issues)}건 이슈")
        return issues

    def _parse_header(self, header_row: list) -> Dict[str, int]:
        """
        헤더 행에서 {헤더 텍스트: 열 인덱스} 매핑을 반환합니다.
        """
        result = {}
        for col_idx, cell in enumerate(header_row):
            text = str(cell.get('value', '')).strip()
            if text:
                result[text] = col_idx
        return result
