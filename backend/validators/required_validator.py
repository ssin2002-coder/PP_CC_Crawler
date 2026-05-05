"""
필수 항목 검증기
지정된 레이블(정산일, 설비, 부서 등)에 대응하는 셀이 비어 있지 않은지 검증합니다.
레이블 텍스트를 찾은 후 인접 셀을 확인합니다.
"""

import logging
import uuid
from typing import Dict, List, Optional, Tuple

from backend.validators.base import BaseValidator, Issue

logger = logging.getLogger(__name__)

# 기본 필수 항목 레이블
DEFAULT_REQUIRED_FIELDS = ('정산일', '설비', '부서')


class RequiredValidator(BaseValidator):
    """
    필수 항목 누락 여부를 확인하는 검증기.

    레이블 텍스트를 셀에서 검색하고, 인접한 셀(오른쪽 또는 아래)이
    비어 있으면 이슈를 생성합니다.
    """

    def validate(
        self,
        cells: list,
        rule_config: dict,
        context: dict = None,
    ) -> List[Issue]:
        """
        필수 항목 검증을 수행합니다.

        rule_config 주요 키:
            fields (list): 검사할 레이블 텍스트 목록
            search_adjacent (str): 'right' 또는 'below' (기본 'right')

        Returns:
            List[Issue]: 누락된 필수 항목 이슈 목록
        """
        issues: List[Issue] = []
        fields = rule_config.get('fields', list(DEFAULT_REQUIRED_FIELDS))
        search_adjacent = rule_config.get('search_adjacent', 'right')
        rule_id = rule_config.get('id', 'auto_required')
        rule_name = rule_config.get('name', '필수 항목 누락')
        severity = rule_config.get('severity', 'error')
        sheet = rule_config.get('sheet', '')

        if not cells:
            return issues

        # 각 필수 레이블 검색
        for field_label in fields:
            position = self._find_label(cells, field_label)
            if position is None:
                # 레이블 자체가 없는 경우
                issues.append(Issue(
                    id=str(uuid.uuid4()),
                    rule_id=rule_id,
                    rule_name=rule_name,
                    severity=severity,
                    cell_ref='N/A',
                    sheet=sheet,
                    message=f"필수 레이블 없음: '{field_label}' 항목을 찾을 수 없습니다.",
                    current_value=None,
                    expected_value=field_label,
                ))
                continue

            row_idx, col_idx = position
            value_cell = self._get_adjacent_cell(cells, row_idx, col_idx, search_adjacent)

            if value_cell is None or self._is_empty_cell(value_cell):
                # 레이블은 있지만 값이 비어 있는 경우
                label_ref = self._cell_ref(row_idx + 1, col_idx + 1)
                if value_cell:
                    # 인접 셀 참조 계산
                    adj_row = row_idx if search_adjacent == 'right' else row_idx + 1
                    adj_col = col_idx + 1 if search_adjacent == 'right' else col_idx
                    value_ref = self._cell_ref(adj_row + 1, adj_col + 1)
                else:
                    value_ref = 'N/A'

                issues.append(Issue(
                    id=str(uuid.uuid4()),
                    rule_id=rule_id,
                    rule_name=rule_name,
                    severity=severity,
                    cell_ref=value_ref,
                    sheet=sheet,
                    message=f"필수 항목 누락: '{field_label}' 값이 비어 있습니다. (레이블: {label_ref})",
                    current_value='(빈 값)',
                    expected_value=f"{field_label} 값 필요",
                ))

        logger.debug(f"필수 항목 검증 완료: {len(issues)}건 이슈")
        return issues

    def _find_label(
        self,
        cells: list,
        label: str,
    ) -> Optional[Tuple[int, int]]:
        """
        셀 매트릭스에서 레이블 텍스트를 포함하는 셀의 위치를 찾습니다.

        Args:
            cells: 2D 셀 매트릭스
            label: 검색할 레이블 텍스트

        Returns:
            Tuple[int, int] | None: (행 인덱스, 열 인덱스), 없으면 None
        """
        for row_idx, row in enumerate(cells):
            for col_idx, cell in enumerate(row):
                cell_text = str(cell.get('value', '')).strip()
                if label in cell_text:
                    return (row_idx, col_idx)
        return None

    def _get_adjacent_cell(
        self,
        cells: list,
        row_idx: int,
        col_idx: int,
        direction: str,
    ) -> Optional[Dict]:
        """
        지정된 방향의 인접 셀을 반환합니다.

        Args:
            cells: 2D 셀 매트릭스
            row_idx: 현재 행 인덱스
            col_idx: 현재 열 인덱스
            direction: 'right' 또는 'below'

        Returns:
            dict | None: 인접 셀 딕셔너리, 없으면 None
        """
        if direction == 'right':
            next_col = col_idx + 1
            if next_col < len(cells[row_idx]):
                return cells[row_idx][next_col]
        elif direction == 'below':
            next_row = row_idx + 1
            if next_row < len(cells) and col_idx < len(cells[next_row]):
                return cells[next_row][col_idx]
        return None
