"""
이상치 검증기
숫자 열에서 Z-score 또는 IQR 방법으로 이상치를 탐지합니다.
이력 통계를 컨텍스트로 제공하면 이력 기반 비교도 수행합니다.
"""

import logging
import math
import uuid
from typing import Dict, List, Optional, Tuple

from backend.validators.base import BaseValidator, Issue

logger = logging.getLogger(__name__)

# 단가/금액 관련 열 헤더 키워드
PRICE_COLUMN_KEYWORDS = ('단가', '금액', '가격', '비용', '단위가격')


class OutlierValidator(BaseValidator):
    """
    Z-score 또는 IQR 방법으로 숫자 열의 이상치를 탐지하는 검증기.

    Context에서 이력 통계 (mean, std)를 제공하면
    현재 데이터 통계 대신 이력 통계를 사용합니다.
    """

    def validate(
        self,
        cells: list,
        rule_config: dict,
        context: dict = None,
    ) -> List[Issue]:
        """
        이상치 검증을 수행합니다.

        rule_config 주요 키:
            method (str): 'zscore' 또는 'iqr' (기본 'zscore')
            threshold (float): Z-score 임계값 (기본 3.0)
            iqr_multiplier (float): IQR 배수 (기본 1.5)
            check_columns (list): 검사할 열 인덱스 목록. 비어 있으면 자동 탐지.

        Returns:
            List[Issue]: 이상치 이슈 목록
        """
        issues: List[Issue] = []
        method = rule_config.get('method', 'zscore')
        threshold = float(rule_config.get('threshold', 3.0))
        iqr_multiplier = float(rule_config.get('iqr_multiplier', 1.5))
        check_columns = rule_config.get('check_columns', [])
        rule_id = rule_config.get('id', 'auto_outlier')
        rule_name = rule_config.get('name', '이상치 탐지')
        severity = rule_config.get('severity', 'warning')
        sheet = rule_config.get('sheet', '')

        if not cells:
            return issues

        # 검사할 열 결정
        target_cols = self._determine_target_columns(cells, check_columns)

        for col_idx in target_cols:
            col_issues = self._check_column(
                cells=cells,
                col_idx=col_idx,
                method=method,
                threshold=threshold,
                iqr_multiplier=iqr_multiplier,
                context=context,
                rule_id=rule_id,
                rule_name=rule_name,
                severity=severity,
                sheet=sheet,
            )
            issues.extend(col_issues)

        logger.debug(f"이상치 검증 완료: {len(issues)}건 이슈")
        return issues

    def _determine_target_columns(
        self,
        cells: list,
        check_columns: list,
    ) -> List[int]:
        """
        검사할 열 인덱스 목록을 결정합니다.
        check_columns가 비어 있으면 헤더에서 자동 탐지합니다.
        """
        if check_columns:
            return [int(c) for c in check_columns]

        # 첫 번째 행을 헤더로 간주하여 자동 탐지
        if not cells:
            return []

        header_row = cells[0]
        target = []
        for col_idx, cell in enumerate(header_row):
            header_text = str(cell.get('value', '')).strip()
            if any(kw in header_text for kw in PRICE_COLUMN_KEYWORDS):
                target.append(col_idx)

        # 자동 탐지 실패 시 모든 숫자 열 검사
        if not target:
            target = self._find_all_numeric_columns(cells)

        return target

    def _find_all_numeric_columns(self, cells: list) -> List[int]:
        """데이터 행에서 숫자 값이 많은 열을 찾습니다."""
        if len(cells) < 2:
            return []

        col_count = max(len(row) for row in cells)
        numeric_cols = []
        for col_idx in range(1, col_count):  # 첫 번째 열(항목명)은 제외
            numeric_count = sum(
                1 for row in cells[1:]
                if col_idx < len(row) and self._get_numeric_value(row[col_idx]) is not None
            )
            if numeric_count >= 2:
                numeric_cols.append(col_idx)
        return numeric_cols

    def _check_column(
        self,
        cells: list,
        col_idx: int,
        method: str,
        threshold: float,
        iqr_multiplier: float,
        context: Optional[dict],
        rule_id: str,
        rule_name: str,
        severity: str,
        sheet: str,
    ) -> List[Issue]:
        """단일 열에 대해 이상치 검사를 수행합니다."""
        # 데이터 행에서 숫자 값 수집 (헤더 제외)
        values_with_pos: List[Tuple[float, int, int]] = []  # (value, row_idx, col_idx)
        for row_idx, row in enumerate(cells[1:], start=1):
            if col_idx >= len(row):
                continue
            cell = row[col_idx]
            val = self._get_numeric_value(cell)
            if val is not None and val != 0:
                values_with_pos.append((val, row_idx, col_idx))

        if len(values_with_pos) < 2:
            return []

        numeric_values = [v[0] for v in values_with_pos]

        if method == 'iqr':
            outlier_indices = self._detect_iqr(numeric_values, iqr_multiplier)
        else:
            # 이력 통계가 있으면 이력 기반으로, 없으면 현재 데이터 기반으로
            hist_mean = None
            hist_std = None
            if context and 'history_stats' in context:
                stats = context['history_stats']
                hist_mean = stats.get('mean')
                hist_std = stats.get('std')
            outlier_indices = self._detect_zscore(numeric_values, threshold, hist_mean, hist_std)

        issues = []
        for idx in outlier_indices:
            val, row_idx, c_idx = values_with_pos[idx]
            cell_ref = self._cell_ref(row_idx + 1, c_idx + 1)
            z_score = self._compute_zscore(val, numeric_values)

            issues.append(Issue(
                id=str(uuid.uuid4()),
                rule_id=rule_id,
                rule_name=rule_name,
                severity=severity,
                cell_ref=cell_ref,
                sheet=sheet,
                message=(
                    f"이상치 감지: {val:,.2f} "
                    f"(Z-score: {z_score:.2f}, 임계값: ±{threshold})"
                ),
                current_value=str(val),
                expected_value=f"범위 내 (±{threshold}σ)",
            ))

        return issues

    def _detect_zscore(
        self,
        values: List[float],
        threshold: float,
        hist_mean: Optional[float] = None,
        hist_std: Optional[float] = None,
    ) -> List[int]:
        """Z-score 방법으로 이상치 인덱스를 반환합니다."""
        if hist_mean is not None and hist_std is not None:
            mean = hist_mean
            std = hist_std
        else:
            mean = sum(values) / len(values)
            variance = sum((v - mean) ** 2 for v in values) / len(values)
            std = math.sqrt(variance)

        if std == 0:
            return []

        return [
            i for i, v in enumerate(values)
            if abs((v - mean) / std) > threshold
        ]

    def _detect_iqr(self, values: List[float], multiplier: float) -> List[int]:
        """IQR 방법으로 이상치 인덱스를 반환합니다."""
        sorted_vals = sorted(values)
        n = len(sorted_vals)
        q1 = sorted_vals[n // 4]
        q3 = sorted_vals[(3 * n) // 4]
        iqr = q3 - q1

        if iqr == 0:
            return []

        lower = q1 - multiplier * iqr
        upper = q3 + multiplier * iqr

        return [
            i for i, v in enumerate(values)
            if v < lower or v > upper
        ]

    def _compute_zscore(self, value: float, values: List[float]) -> float:
        """단일 값의 Z-score를 계산합니다."""
        if len(values) < 2:
            return 0.0
        mean = sum(values) / len(values)
        variance = sum((v - mean) ** 2 for v in values) / len(values)
        std = math.sqrt(variance)
        if std == 0:
            return 0.0
        return (value - mean) / std
