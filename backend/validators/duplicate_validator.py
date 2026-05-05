"""
중복 항목 검증기
동일 항목이 같은 금액으로 2회 이상 등장하는 경우를 탐지합니다.
"""

import logging
import uuid
from collections import defaultdict
from typing import Dict, List, Tuple

from backend.validators.base import BaseValidator, Issue

logger = logging.getLogger(__name__)


class DuplicateValidator(BaseValidator):
    """
    중복 항목을 탐지하는 검증기.

    키 열(항목명)과 값 열(금액)의 조합이 중복되는 경우를 찾습니다.
    """

    def validate(
        self,
        cells: list,
        rule_config: dict,
        context: dict = None,
    ) -> List[Issue]:
        """
        중복 항목 검증을 수행합니다.

        rule_config 주요 키:
            key_columns (list): 키 열 인덱스 목록 (기본 [0])
            value_columns (list): 값 열 인덱스 목록 (기본 [4])
            ignore_empty (bool): 빈 키는 무시 여부 (기본 True)
            min_occurrences (int): 이슈 발생 최소 중복 횟수 (기본 2)

        Returns:
            List[Issue]: 중복 이슈 목록
        """
        issues: List[Issue] = []
        key_columns: List[int] = rule_config.get('key_columns', [0])
        value_columns: List[int] = rule_config.get('value_columns', [4])
        ignore_empty: bool = rule_config.get('ignore_empty', True)
        min_occurrences: int = int(rule_config.get('min_occurrences', 2))
        rule_id = rule_config.get('id', 'auto_duplicate')
        rule_name = rule_config.get('name', '중복 항목 검출')
        severity = rule_config.get('severity', 'warning')
        sheet = rule_config.get('sheet', '')

        if not cells or len(cells) < 2:
            return issues

        # (키, 값) -> [(행 인덱스, 열 인덱스)] 매핑
        occurrence_map: Dict[Tuple, List[Tuple[int, int]]] = defaultdict(list)

        # 헤더 행(0번)은 제외하고 데이터 행 순회
        for row_idx, row in enumerate(cells[1:], start=1):
            # 키 값 추출
            key_parts = []
            for k_col in key_columns:
                if k_col < len(row):
                    key_parts.append(str(row[k_col].get('value', '')).strip())
                else:
                    key_parts.append('')

            key = tuple(key_parts)

            # 빈 키 무시
            if ignore_empty and all(k == '' for k in key):
                continue

            # 소계/합계 행 제외
            first_key = key_parts[0] if key_parts else ''
            if any(kw in first_key for kw in ('소계', '합계', '계', 'TOTAL')):
                continue

            # 값 열 추출
            value_parts = []
            for v_col in value_columns:
                if v_col < len(row):
                    val = self._get_numeric_value(row[v_col])
                    value_parts.append(val)
                else:
                    value_parts.append(None)

            composite_key = (key, tuple(value_parts))

            # 첫 번째 값 열 인덱스를 대표 열로 사용
            ref_col = value_columns[0] if value_columns else key_columns[0]
            occurrence_map[composite_key].append((row_idx, ref_col))

        # 중복 발생한 항목에 대해 이슈 생성
        for (key, values), positions in occurrence_map.items():
            if len(positions) < min_occurrences:
                continue

            key_str = ' | '.join(key)
            value_str = ', '.join(
                f"{v:,.0f}" if isinstance(v, float) else str(v)
                for v in values
            )

            for pos_idx, (row_idx, col_idx) in enumerate(positions):
                cell_ref = self._cell_ref(row_idx + 1, col_idx + 1)
                issues.append(Issue(
                    id=str(uuid.uuid4()),
                    rule_id=rule_id,
                    rule_name=rule_name,
                    severity=severity,
                    cell_ref=cell_ref,
                    sheet=sheet,
                    message=(
                        f"중복 항목: '{key_str}' (금액: {value_str}) — "
                        f"총 {len(positions)}회 중복 ({pos_idx + 1}번째)"
                    ),
                    current_value=key_str,
                    expected_value=f"1회 이하 (현재: {len(positions)}회)",
                ))

        logger.debug(f"중복 검증 완료: {len(issues)}건 이슈")
        return issues
