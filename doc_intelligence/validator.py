"""
validator.py — 6종 룰 + 교차 검증기
BaseRule: ValueMatchRule, OrderCheckRule, FormulaCheckRule,
          ExistsRule, ContainsRule, RangeCheckRule
CrossValidator: 룰 조합 실행, Engine 플러그인 인터페이스 구현
"""
from datetime import datetime


# ──────────────────────────────────────────────
# 공통 유틸
# ──────────────────────────────────────────────

_DATE_FORMATS = ["%Y.%m.%d", "%Y-%m-%d", "%Y/%m/%d"]


def _parse_date(value: str) -> datetime:
    """지원 형식(%Y.%m.%d / %Y-%m-%d / %Y/%m/%d)으로 날짜 파싱. 실패 시 ValueError."""
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(value.strip(), fmt)
        except ValueError:
            continue
    raise ValueError(f"날짜 파싱 실패: {value!r}")


# ──────────────────────────────────────────────
# BaseRule
# ──────────────────────────────────────────────

class BaseRule:
    def __init__(self, name: str):
        self.name = name

    def check(self) -> dict:
        raise NotImplementedError


# ──────────────────────────────────────────────
# 1. ValueMatchRule
# ──────────────────────────────────────────────

class ValueMatchRule(BaseRule):
    """regions 내 모든 value가 동일하면 '통과', 아니면 '실패'."""

    def __init__(self, name: str, regions: list):
        super().__init__(name)
        self.regions = regions  # [{"value": "100"}, {"value": "100"}, ...]

    def check(self) -> dict:
        values = [r["value"] for r in self.regions]
        if len(set(values)) <= 1:
            return {"rule": self.name, "status": "통과", "detail": f"모든 값 일치: {values[0]!r}"}
        return {
            "rule": self.name,
            "status": "실패",
            "detail": f"값 불일치: {values}",
        }


# ──────────────────────────────────────────────
# 2. OrderCheckRule
# ──────────────────────────────────────────────

class OrderCheckRule(BaseRule):
    """regions 날짜 값이 오름차순이면 '통과', 역순(내림차순 포함)이면 '실패'."""

    def __init__(self, name: str, regions: list):
        super().__init__(name)
        self.regions = regions  # [{"value": "2025.01.01"}, {"value": "2025.02.01"}]

    def check(self) -> dict:
        try:
            dates = [_parse_date(r["value"]) for r in self.regions]
        except ValueError as exc:
            return {"rule": self.name, "status": "실패", "detail": str(exc)}

        for i in range(len(dates) - 1):
            if dates[i] > dates[i + 1]:
                return {
                    "rule": self.name,
                    "status": "실패",
                    "detail": f"날짜 역순: {self.regions[i]['value']} > {self.regions[i + 1]['value']}",
                }
        return {
            "rule": self.name,
            "status": "통과",
            "detail": "날짜 순서 정상",
        }


# ──────────────────────────────────────────────
# 3. FormulaCheckRule
# ──────────────────────────────────────────────

class FormulaCheckRule(BaseRule):
    """
    operands와 operator로 계산값을 구해 expected와 비교.
    일치: '통과', 차이 <=10: '경고', 그 외: '실패'.
    operator: '*' | '+' | '-' | '/'
    """

    def __init__(self, name: str, params: dict):
        super().__init__(name)
        # params = {
        #   "operands": [{"value": "10"}, {"value": "1500000"}],
        #   "operator": "*",
        #   "expected": {"value": "15000000"}
        # }
        self.params = params

    def check(self) -> dict:
        try:
            operands = [float(o["value"]) for o in self.params["operands"]]
            expected = float(self.params["expected"]["value"])
            operator = self.params["operator"]
        except (KeyError, ValueError) as exc:
            return {"rule": self.name, "status": "실패", "detail": f"파라미터 오류: {exc}"}

        try:
            calculated = operands[0]
            for operand in operands[1:]:
                if operator == "*":
                    calculated *= operand
                elif operator == "+":
                    calculated += operand
                elif operator == "-":
                    calculated -= operand
                elif operator == "/":
                    calculated /= operand
                else:
                    return {"rule": self.name, "status": "실패", "detail": f"지원하지 않는 연산자: {operator!r}"}
        except ZeroDivisionError:
            return {"rule": self.name, "status": "실패", "detail": "0으로 나누기 오류"}

        diff = abs(calculated - expected)
        if diff == 0:
            return {
                "rule": self.name,
                "status": "통과",
                "detail": f"계산값 {calculated} == 기대값 {expected}",
            }
        if diff <= 10:
            return {
                "rule": self.name,
                "status": "경고",
                "detail": f"계산값 {calculated}과 기대값 {expected}의 차이 {diff} (허용 범위 내)",
            }
        return {
            "rule": self.name,
            "status": "실패",
            "detail": f"계산값 {calculated} != 기대값 {expected} (차이: {diff})",
        }


# ──────────────────────────────────────────────
# 4. ExistsRule
# ──────────────────────────────────────────────

class ExistsRule(BaseRule):
    """regions 내 빈 값이 없으면 '통과', 있으면 '실패'."""

    def __init__(self, name: str, regions: list):
        super().__init__(name)
        self.regions = regions  # [{"value": "abc"}, {"value": ""}]

    def check(self) -> dict:
        empty = [i for i, r in enumerate(self.regions) if not r["value"]]
        if not empty:
            return {"rule": self.name, "status": "통과", "detail": "모든 값 존재"}
        return {
            "rule": self.name,
            "status": "실패",
            "detail": f"빈 값 위치(인덱스): {empty}",
        }


# ──────────────────────────────────────────────
# 5. ContainsRule
# ──────────────────────────────────────────────

class ContainsRule(BaseRule):
    """source['value']가 target['value']에 포함되면 '통과', 아니면 '실패'."""

    def __init__(self, name: str, source: dict, target: dict):
        super().__init__(name)
        self.source = source  # {"value": "..."}
        self.target = target  # {"value": "..."}

    def check(self) -> dict:
        src = self.source["value"]
        tgt = self.target["value"]
        if src in tgt:
            return {
                "rule": self.name,
                "status": "통과",
                "detail": f"{src!r}이 {tgt!r}에 포함됨",
            }
        return {
            "rule": self.name,
            "status": "실패",
            "detail": f"{src!r}이 {tgt!r}에 포함되지 않음",
        }


# ──────────────────────────────────────────────
# 6. RangeCheckRule
# ──────────────────────────────────────────────

class RangeCheckRule(BaseRule):
    """
    target 값이 bounds의 min~max 범위 내에 있으면 '통과', 아니면 '실패'.
    날짜 파싱 먼저 시도, 실패 시 float으로 처리.
    """

    def __init__(self, name: str, target: dict, bounds: dict):
        super().__init__(name)
        self.target = target   # {"value": "2025.02.15"}
        self.bounds = bounds   # {"min": "2025.01.01", "max": "2025.03.31"}

    def check(self) -> dict:
        target_val = self.target["value"]
        min_val = self.bounds["min"]
        max_val = self.bounds["max"]

        # 날짜로 파싱 시도
        try:
            t = _parse_date(target_val)
            lo = _parse_date(min_val)
            hi = _parse_date(max_val)
            if lo <= t <= hi:
                return {
                    "rule": self.name,
                    "status": "통과",
                    "detail": f"{target_val}이 [{min_val}, {max_val}] 범위 내",
                }
            return {
                "rule": self.name,
                "status": "실패",
                "detail": f"{target_val}이 [{min_val}, {max_val}] 범위 밖",
            }
        except ValueError:
            pass

        # 숫자로 파싱 시도
        try:
            t_num = float(target_val)
            lo_num = float(min_val)
            hi_num = float(max_val)
        except ValueError as exc:
            return {"rule": self.name, "status": "실패", "detail": f"값 파싱 실패: {exc}"}

        if lo_num <= t_num <= hi_num:
            return {
                "rule": self.name,
                "status": "통과",
                "detail": f"{target_val}이 [{min_val}, {max_val}] 범위 내",
            }
        return {
            "rule": self.name,
            "status": "실패",
            "detail": f"{target_val}이 [{min_val}, {max_val}] 범위 밖",
        }


# ──────────────────────────────────────────────
# CrossValidator (Engine 플러그인 인터페이스)
# ──────────────────────────────────────────────

class CrossValidator:
    """등록된 룰 전체를 실행하고 결과를 context에 저장하는 Engine 플러그인."""

    name = "validator"
    enabled = True

    def __init__(self):
        self.rules: list = []

    def initialize(self, engine) -> None:
        """Engine.register() 시 호출. 현재는 별도 초기화 없음."""
        pass

    def add_rule(self, rule: BaseRule) -> None:
        """룰 추가."""
        self.rules.append(rule)

    def validate(self) -> list:
        """등록된 모든 룰의 check() 결과를 리스트로 반환."""
        return [rule.check() for rule in self.rules]

    def process(self, doc, context: dict) -> dict:
        """Engine 파이프라인 인터페이스. validate() 실행 후 context에 저장."""
        context["validation_results"] = self.validate()
        return context
