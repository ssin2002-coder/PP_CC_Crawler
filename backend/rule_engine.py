"""
규칙 엔진 모듈
활성화된 규칙을 로드하고 적절한 검증기를 디스패치하여
모든 이슈를 수집·정렬합니다.
"""

import logging
import uuid
from typing import Dict, List, Optional, Type

from backend.validators.base import BaseValidator, Issue

logger = logging.getLogger(__name__)

# 심각도 정렬 우선순위
SEVERITY_ORDER = {'error': 0, 'warning': 1, 'info': 2}


class RuleEngine:
    """
    검증 규칙을 관리하고 검증기를 조율하는 엔진 클래스.

    storage에서 활성 규칙을 로드하고, 각 규칙의 template에 맞는
    검증기를 찾아 실행한 후 결과를 집계합니다.
    """

    def __init__(self, storage, validators_map: Dict[str, Type[BaseValidator]]) -> None:
        """
        Args:
            storage: storage 모듈 (load_rules 함수 제공)
            validators_map: {'template_name': ValidatorClass} 매핑
        """
        self._storage = storage
        self._validators_map = validators_map

    def run_validation(
        self,
        cells: list,
        sheet: str = '',
        rules: Optional[List[Dict]] = None,
        context: Optional[Dict] = None,
    ) -> Dict:
        """
        전체 검증을 실행하고 결과를 반환합니다.

        Args:
            cells: 2D 셀 매트릭스 (read_range 결과)
            sheet: 시트 이름 (이슈에 포함)
            rules: 사용할 규칙 목록. None이면 storage에서 로드.
            context: 추가 컨텍스트 (이력 통계 등)

        Returns:
            dict: {
                'issues': List[dict],
                'summary': {'errors': int, 'warnings': int, 'info': int, 'ok': bool},
                'total': int,
            }
        """
        # 규칙 로드
        if rules is None:
            rules = self._storage.load_rules()

        # 활성화된 규칙만 필터
        active_rules = [r for r in rules if r.get('enabled', True)]
        logger.info(f"검증 시작: 활성 규칙 {len(active_rules)}개, 시트: '{sheet}'")

        all_issues: List[Issue] = []

        for rule in active_rules:
            rule_issues = self._run_single_rule(
                rule=rule,
                cells=cells,
                sheet=sheet,
                context=context,
            )
            all_issues.extend(rule_issues)

        # 심각도 기준 정렬 (error > warning > info)
        all_issues.sort(key=lambda i: SEVERITY_ORDER.get(i.severity, 99))

        # 고유 ID 재부여 (순서 기반)
        for idx, issue in enumerate(all_issues):
            if not issue.id:
                issue.id = str(uuid.uuid4())

        # 요약 통계
        summary = self._build_summary(all_issues)

        logger.info(
            f"검증 완료: 오류 {summary['errors']}건, "
            f"경고 {summary['warnings']}건, "
            f"정보 {summary['info']}건"
        )

        return {
            'issues': [issue.to_dict() for issue in all_issues],
            'summary': summary,
            'total': len(all_issues),
        }

    def _run_single_rule(
        self,
        rule: Dict,
        cells: list,
        sheet: str,
        context: Optional[Dict],
    ) -> List[Issue]:
        """
        단일 규칙에 대한 검증을 실행합니다.

        Args:
            rule: 규칙 딕셔너리
            cells: 셀 매트릭스
            sheet: 시트 이름
            context: 컨텍스트 데이터

        Returns:
            List[Issue]: 이 규칙에서 발견된 이슈 목록
        """
        template = rule.get('template')
        if not template:
            logger.warning(f"규칙에 template이 없음: {rule.get('id')}")
            return []

        validator_class = self._validators_map.get(template)
        if validator_class is None:
            logger.warning(f"알 수 없는 template: {template} (규칙: {rule.get('id')})")
            return []

        # 규칙 config에 시트 정보와 규칙 메타 추가
        rule_config = {
            **rule.get('config', {}),
            'id': rule.get('id', ''),
            'name': rule.get('name', ''),
            'severity': rule.get('severity', 'warning'),
            'sheet': sheet,
        }

        try:
            validator = validator_class()
            issues = validator.validate(cells, rule_config, context)
            logger.debug(f"규칙 '{rule.get('name')}' 완료: {len(issues)}건 이슈")
            return issues
        except Exception as e:
            logger.error(f"규칙 실행 오류 [{rule.get('id')}]: {e}", exc_info=True)
            return []

    def _build_summary(self, issues: List[Issue]) -> Dict:
        """
        이슈 목록에서 요약 통계를 생성합니다.

        Returns:
            dict: {'errors': int, 'warnings': int, 'info': int, 'ok': bool}
        """
        errors = sum(1 for i in issues if i.severity == 'error')
        warnings = sum(1 for i in issues if i.severity == 'warning')
        info = sum(1 for i in issues if i.severity == 'info')

        return {
            'errors': errors,
            'warnings': warnings,
            'info': info,
            'ok': errors == 0 and warnings == 0,
        }
