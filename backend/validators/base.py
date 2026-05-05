"""
검증기 기반 클래스 및 Issue 데이터 클래스
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, asdict
from typing import List, Optional


@dataclass
class Issue:
    """검증에서 발견된 단일 이슈를 나타냅니다."""

    id: str
    """이슈 고유 ID (UUID 형식)"""

    rule_id: str
    """이 이슈를 생성한 규칙의 ID"""

    rule_name: str
    """이 이슈를 생성한 규칙의 이름"""

    severity: str
    """심각도: 'error', 'warning', 'info' 중 하나"""

    cell_ref: str
    """이슈가 발생한 셀 참조 ('A1' 형식)"""

    sheet: str
    """이슈가 발생한 시트 이름"""

    message: str
    """사용자에게 표시할 이슈 메시지"""

    current_value: Optional[str] = None
    """현재 셀 값 (있는 경우)"""

    expected_value: Optional[str] = None
    """기대하는 값 (있는 경우)"""

    def to_dict(self) -> dict:
        """Issue를 직렬화 가능한 딕셔너리로 변환합니다."""
        return asdict(self)


class BaseValidator(ABC):
    """
    모든 검증기의 기반 추상 클래스.

    서브클래스는 validate() 메서드를 구현해야 합니다.
    """

    @abstractmethod
    def validate(
        self,
        cells: list,
        rule_config: dict,
        context: dict = None,
    ) -> List[Issue]:
        """
        셀 데이터를 검증하고 발견된 이슈 목록을 반환합니다.

        Args:
            cells: 2D 매트릭스 형태의 셀 딕셔너리 목록.
                   각 셀: {value, raw_value, row, col, is_bold, is_number, bg_color, font_color}
            rule_config: storage에서 가져온 규칙 설정 JSON
            context: 추가 컨텍스트 데이터 (이력 통계 등)

        Returns:
            List[Issue]: 발견된 이슈 목록
        """
        pass

    # ------------------------------------------------------------------
    # 공통 유틸리티 메서드
    # ------------------------------------------------------------------

    def _col_letter(self, col: int) -> str:
        """
        1-indexed 열 번호를 Excel 열 문자로 변환합니다.
        예: 1 -> 'A', 26 -> 'Z', 27 -> 'AA'
        """
        result = ''
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _cell_ref(self, row: int, col: int) -> str:
        """
        행/열 번호(1-indexed)를 셀 참조 문자열로 변환합니다.
        예: row=1, col=1 -> 'A1'
        """
        return f"{self._col_letter(col)}{row}"

    def _get_numeric_value(self, cell: dict) -> Optional[float]:
        """
        셀에서 숫자 값을 추출합니다.
        raw_value 우선, 없으면 value 필드에서 숫자 파싱을 시도합니다.
        숫자가 아닌 셀이면 None을 반환합니다.
        """
        import re
        # raw_value 우선 시도
        raw = cell.get('raw_value')
        if raw is not None:
            try:
                return float(raw)
            except (TypeError, ValueError):
                pass

        # value 필드에서 숫자 파싱 (콤마 제거)
        value = cell.get('value', '')
        if value is None or str(value).strip() == '':
            return None
        cleaned = re.sub(r'[,\s]', '', str(value))
        try:
            return float(cleaned)
        except (TypeError, ValueError):
            return None

    def _is_empty_cell(self, cell: dict) -> bool:
        """셀이 비어 있는지 확인합니다."""
        value = cell.get('value', '')
        raw = cell.get('raw_value')
        # raw_value가 명시적으로 None이고 value도 비어있으면 빈 셀
        if raw is not None:
            return str(raw).strip() == ''
        return str(value).strip() == ''
