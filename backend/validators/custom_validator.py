"""
사용자 정의 검증기
JSON으로 기술된 IF/AND/OR/THEN 조건 블록 규칙을 평가합니다.
"""

import logging
import uuid
from typing import Any, Dict, List, Optional

from backend.validators.base import BaseValidator, Issue

logger = logging.getLogger(__name__)

# 지원 연산자
SUPPORTED_OPERATORS = ('==', '!=', '>', '<', '>=', '<=', 'contains', 'is_empty')

# 지원 집계 함수
SUPPORTED_AGGREGATES = ('sum', 'avg', 'max', 'min', 'count')

# 액션 타입
ACTION_SEVERITY_MAP = {
    'mark_error': 'error',
    'mark_warning': 'warning',
    'mark_info': 'info',
}


class CustomValidator(BaseValidator):
    """
    JSON 기반 IF/THEN 규칙을 평가하는 사용자 정의 검증기.

    규칙 구조 예시:
    {
        "conditions": [
            {
                "type": "AND",
                "rules": [
                    {"left": "A1", "op": ">", "right": 0},
                    {"left": "col:sum:4", "op": ">", "right": "col:sum:3"}
                ]
            }
        ],
        "action": {
            "type": "mark_error",
            "cell": "A1",
            "message": "금액이 합계를 초과합니다."
        }
    }
    """

    def validate(
        self,
        cells: list,
        rule_config: dict,
        context: dict = None,
    ) -> List[Issue]:
        """
        사용자 정의 규칙 조건을 평가합니다.

        rule_config 주요 키:
            conditions (list): 조건 블록 목록 (AND/OR 구조)
            action (dict): 조건 충족 시 수행할 액션
                - type: 'mark_error' | 'mark_warning' | 'mark_info'
                - cell: 이슈 셀 참조 (예: "A1", 또는 "auto")
                - message: 이슈 메시지

        Returns:
            List[Issue]: 조건 충족 시 생성된 이슈 목록
        """
        issues: List[Issue] = []
        conditions = rule_config.get('conditions', [])
        action = rule_config.get('action', {})
        rule_id = rule_config.get('id', 'custom')
        rule_name = rule_config.get('name', '사용자 정의 규칙')
        sheet = rule_config.get('sheet', '')

        if not cells or not conditions or not action:
            return issues

        # 컨텍스트 빌드 (집계 값 사전 계산)
        eval_context = self._build_context(cells)

        # 각 행에 대해 조건 평가
        for row_idx, row in enumerate(cells):
            row_context = {**eval_context, 'current_row': row_idx, 'current_row_data': row}

            if self._evaluate_conditions(conditions, cells, row_context):
                action_type = action.get('type', 'mark_warning')
                severity = ACTION_SEVERITY_MAP.get(action_type, 'warning')
                message = action.get('message', '사용자 정의 조건 위반')
                cell_ref = self._resolve_cell_ref(action.get('cell', 'auto'), row_idx, cells)

                issues.append(Issue(
                    id=str(uuid.uuid4()),
                    rule_id=rule_id,
                    rule_name=rule_name,
                    severity=severity,
                    cell_ref=cell_ref,
                    sheet=sheet,
                    message=message,
                    current_value=None,
                    expected_value=None,
                ))

        logger.debug(f"사용자 정의 검증 완료: {len(issues)}건 이슈")
        return issues

    def _build_context(self, cells: list) -> Dict[str, Any]:
        """
        집계 값을 사전 계산하여 평가 컨텍스트를 빌드합니다.
        col:sum:N, col:avg:N 등의 참조를 위한 캐시입니다.
        """
        context: Dict[str, Any] = {}

        if not cells:
            return context

        col_count = max((len(row) for row in cells), default=0)

        for col_idx in range(col_count):
            values = []
            for row in cells[1:]:  # 헤더 제외
                if col_idx < len(row):
                    val = self._get_numeric_value(row[col_idx])
                    if val is not None:
                        values.append(val)

            if values:
                context[f'col:sum:{col_idx}'] = sum(values)
                context[f'col:avg:{col_idx}'] = sum(values) / len(values)
                context[f'col:max:{col_idx}'] = max(values)
                context[f'col:min:{col_idx}'] = min(values)
                context[f'col:count:{col_idx}'] = len(values)
            else:
                context[f'col:sum:{col_idx}'] = 0
                context[f'col:avg:{col_idx}'] = 0
                context[f'col:max:{col_idx}'] = 0
                context[f'col:min:{col_idx}'] = 0
                context[f'col:count:{col_idx}'] = 0

        return context

    def _evaluate_conditions(
        self,
        conditions: list,
        cells: list,
        context: Dict[str, Any],
    ) -> bool:
        """
        조건 블록 목록을 평가합니다.
        최상위 조건 목록은 AND로 연결됩니다.
        """
        for condition in conditions:
            cond_type = condition.get('type', 'AND').upper()
            rules = condition.get('rules', [])

            if cond_type == 'AND':
                if not all(self._evaluate_single_rule(r, cells, context) for r in rules):
                    return False
            elif cond_type == 'OR':
                if rules and not any(self._evaluate_single_rule(r, cells, context) for r in rules):
                    return False
            elif cond_type == 'NOT':
                if rules and self._evaluate_single_rule(rules[0], cells, context):
                    return False

        return True

    def _evaluate_single_rule(
        self,
        rule: Dict,
        cells: list,
        context: Dict[str, Any],
    ) -> bool:
        """
        단일 조건 규칙을 평가합니다.

        rule 구조:
            left: 좌변 (셀 참조, 집계, 리터럴)
            op: 연산자
            right: 우변 (셀 참조, 집계, 리터럴)
        """
        left_raw = rule.get('left')
        op = rule.get('op', '==')
        right_raw = rule.get('right')

        left_val = self._resolve_value(left_raw, cells, context)
        right_val = self._resolve_value(right_raw, cells, context)

        return self._compare(left_val, op, right_val)

    def _resolve_value(
        self,
        ref: Any,
        cells: list,
        context: Dict[str, Any],
    ) -> Any:
        """
        값 참조를 실제 값으로 변환합니다.

        지원 형식:
            - "A1": 셀 참조 (고정)
            - "col:sum:3": 3번 열 합계
            - 숫자/문자열 리터럴
            - "row:col:N": 현재 행의 N번 열
        """
        if ref is None:
            return None

        ref_str = str(ref)

        # 집계 참조 (col:func:N)
        if ref_str.startswith('col:') and ref_str in context:
            return context[ref_str]

        # 현재 행 열 참조 (row:col:N)
        if ref_str.startswith('row:col:'):
            try:
                col_idx = int(ref_str.split(':')[-1])
                row_data = context.get('current_row_data', [])
                if col_idx < len(row_data):
                    return self._get_numeric_value(row_data[col_idx]) or row_data[col_idx].get('value')
            except (ValueError, IndexError):
                pass

        # 셀 참조 (A1 형식)
        if len(ref_str) >= 2 and ref_str[0].isalpha() and ref_str[1:].isdigit():
            row, col = self._parse_cell_ref(ref_str)
            if row is not None and row < len(cells) and col < len(cells[row]):
                cell = cells[row][col]
                return self._get_numeric_value(cell) or cell.get('value')

        # 숫자 리터럴
        try:
            return float(ref)
        except (TypeError, ValueError):
            pass

        # 문자열 리터럴
        return ref_str

    def _compare(self, left: Any, op: str, right: Any) -> bool:
        """두 값을 지정된 연산자로 비교합니다."""
        try:
            if op == 'is_empty':
                return left is None or str(left).strip() == ''

            if op == 'contains':
                return str(right).strip().lower() in str(left).strip().lower()

            # 숫자 비교를 위한 변환 시도
            try:
                l_num = float(left)
                r_num = float(right)
                left, right = l_num, r_num
            except (TypeError, ValueError):
                # 문자열 비교
                left = str(left)
                right = str(right)

            op_map = {
                '==': lambda a, b: a == b,
                '!=': lambda a, b: a != b,
                '>': lambda a, b: a > b,
                '<': lambda a, b: a < b,
                '>=': lambda a, b: a >= b,
                '<=': lambda a, b: a <= b,
            }
            func = op_map.get(op)
            if func is None:
                logger.warning(f"알 수 없는 연산자: {op}")
                return False
            return func(left, right)

        except Exception as e:
            logger.debug(f"비교 오류 ({left} {op} {right}): {e}")
            return False

    def _parse_cell_ref(self, ref: str):
        """
        'A1' 형식의 셀 참조를 0-indexed (row, col) 튜플로 변환합니다.
        """
        try:
            col_str = ''
            row_str = ''
            for ch in ref:
                if ch.isalpha():
                    col_str += ch.upper()
                else:
                    row_str += ch

            # 열 문자를 인덱스로 변환
            col = 0
            for ch in col_str:
                col = col * 26 + (ord(ch) - ord('A') + 1)
            col -= 1  # 0-indexed

            row = int(row_str) - 1  # 0-indexed
            return row, col
        except Exception:
            return None, None

    def _resolve_cell_ref(self, ref: str, row_idx: int, cells: list) -> str:
        """
        액션의 셀 참조를 결정합니다.
        'auto'인 경우 현재 행의 첫 번째 열을 사용합니다.
        """
        if ref == 'auto':
            return self._cell_ref(row_idx + 1, 1)
        return ref
