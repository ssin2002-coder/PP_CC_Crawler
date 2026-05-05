"""
합계 검증기
소계/합계/계 행의 값이 데이터 행의 합과 일치하는지 검증합니다.
자재비, 인건비, 경비 등 여러 섹션을 처리합니다.
"""

import logging
import uuid
from typing import Dict, List, Optional, Tuple

from backend.validators.base import BaseValidator, Issue

logger = logging.getLogger(__name__)

# 소계/합계 행을 식별하는 키워드
SUBTOTAL_KEYWORDS = ('소계', '합계', '계', 'TOTAL', 'Total')


class SumValidator(BaseValidator):
    """
    소계/합계 행의 합산 정확성을 검증하는 검증기.

    첫 번째 열에서 소계/합계 키워드를 찾고,
    해당 행의 숫자 열 값이 위의 데이터 행 합산과 일치하는지 확인합니다.
    """

    def validate(
        self,
        cells: list,
        rule_config: dict,
        context: dict = None,
    ) -> List[Issue]:
        """
        합계 검증을 수행합니다.

        rule_config 주요 키:
            target_section (str): 검사할 섹션 키워드 (예: "자재", "인건"). None이면 전체 검사.
            tolerance (float): 허용 오차 (기본 0).
            severity (str): 이슈 심각도 (기본 "error").

        Returns:
            List[Issue]: 합계 불일치 이슈 목록
        """
        issues: List[Issue] = []
        tolerance = float(rule_config.get('tolerance', 0))
        target_section = rule_config.get('target_section')
        rule_id = rule_config.get('id', 'auto_sum')
        rule_name = rule_config.get('name', '합계 검증')
        severity = rule_config.get('severity', 'error')
        sheet = rule_config.get('sheet', '')

        if not cells:
            return issues

        # 소계 행 위치 탐색
        subtotal_rows = self._find_subtotal_rows(cells, target_section)

        for subtotal_row_idx, amount_col_idx in subtotal_rows:
            issue = self._check_subtotal(
                cells=cells,
                subtotal_row_idx=subtotal_row_idx,
                amount_col_idx=amount_col_idx,
                tolerance=tolerance,
                rule_id=rule_id,
                rule_name=rule_name,
                severity=severity,
                sheet=sheet,
            )
            if issue:
                issues.append(issue)

        logger.debug(f"합계 검증 완료: {len(issues)}건 이슈")
        return issues

    def _find_subtotal_rows(
        self,
        cells: list,
        target_section: Optional[str],
    ) -> List[Tuple[int, int]]:
        """
        소계/합계 행의 인덱스와 금액 열 인덱스를 찾습니다.

        Args:
            cells: 2D 셀 매트릭스
            target_section: 필터링할 섹션 이름 (None이면 전체)

        Returns:
            List[Tuple[int, int]]: (행 인덱스, 금액 열 인덱스) 튜플 목록
        """
        result = []
        in_target_section = target_section is None  # 섹션 필터 없으면 전체

        for row_idx, row in enumerate(cells):
            if not row:
                continue

            # 첫 번째 열 텍스트 확인
            first_cell_value = str(row[0].get('value', '')).strip()

            # 섹션 헤더 감지 (target_section이 지정된 경우)
            if target_section and target_section in first_cell_value:
                in_target_section = True

            if not in_target_section:
                continue

            # 소계/합계 키워드 확인
            is_subtotal = any(kw in first_cell_value for kw in SUBTOTAL_KEYWORDS)
            if not is_subtotal:
                continue

            # 숫자 열(금액 열) 찾기: 첫 번째 열 이후에서 숫자가 있는 마지막 열
            amount_col_idx = self._find_amount_column(row)
            if amount_col_idx is not None:
                result.append((row_idx, amount_col_idx))

        return result

    def _find_amount_column(self, row: list) -> Optional[int]:
        """
        행에서 금액(숫자) 열 인덱스를 반환합니다.
        숫자가 있는 가장 오른쪽 열을 금액 열로 간주합니다.
        """
        amount_col = None
        for col_idx in range(1, len(row)):
            cell = row[col_idx]
            if cell.get('is_number') or self._get_numeric_value(cell) is not None:
                val = self._get_numeric_value(cell)
                if val is not None:
                    amount_col = col_idx
        return amount_col

    def _check_subtotal(
        self,
        cells: list,
        subtotal_row_idx: int,
        amount_col_idx: int,
        tolerance: float,
        rule_id: str,
        rule_name: str,
        severity: str,
        sheet: str,
    ) -> Optional[Issue]:
        """
        소계 행과 위의 데이터 행 합산을 비교합니다.

        Args:
            cells: 2D 셀 매트릭스
            subtotal_row_idx: 소계 행 인덱스 (0-based)
            amount_col_idx: 금액 열 인덱스 (0-based)
            tolerance: 허용 오차
            rule_id, rule_name, severity, sheet: 이슈 메타데이터

        Returns:
            Issue | None: 불일치 시 Issue, 일치 시 None
        """
        subtotal_cell = cells[subtotal_row_idx][amount_col_idx]
        subtotal_value = self._get_numeric_value(subtotal_cell)
        if subtotal_value is None:
            return None

        # 이 소계 행 위의 데이터 행 합산
        computed_sum = 0.0
        data_row_count = 0
        for row_idx in range(subtotal_row_idx - 1, -1, -1):
            row = cells[row_idx]
            if not row or amount_col_idx >= len(row):
                break

            first_value = str(row[0].get('value', '')).strip()
            # 이전 소계/헤더 행이면 중단
            if any(kw in first_value for kw in SUBTOTAL_KEYWORDS):
                break
            # 헤더 행(굵은 글씨)이면 중단
            if row[0].get('is_bold') and data_row_count == 0:
                # 헤더 바로 아래가 데이터이므로 헤더 행 자체는 제외
                continue

            cell = row[amount_col_idx]
            val = self._get_numeric_value(cell)
            if val is not None:
                computed_sum += val
                data_row_count += 1

        if data_row_count == 0:
            return None

        difference = abs(subtotal_value - computed_sum)
        if difference <= tolerance:
            return None

        # 셀 참조 생성 (1-indexed)
        cell_ref = self._cell_ref(subtotal_row_idx + 1, amount_col_idx + 1)

        return Issue(
            id=str(uuid.uuid4()),
            rule_id=rule_id,
            rule_name=rule_name,
            severity=severity,
            cell_ref=cell_ref,
            sheet=sheet,
            message=(
                f"합계 불일치: 표시값 {subtotal_value:,.0f} ≠ "
                f"계산값 {computed_sum:,.0f} "
                f"(차이: {difference:,.0f})"
            ),
            current_value=str(subtotal_value),
            expected_value=str(computed_sum),
        )
