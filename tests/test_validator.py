"""
test_validator.py — 6종 룰 + CrossValidator 테스트 (12개)
"""
import pytest
from doc_intelligence.validator import (
    ValueMatchRule,
    OrderCheckRule,
    FormulaCheckRule,
    ExistsRule,
    ContainsRule,
    RangeCheckRule,
    CrossValidator,
)


# ──────────────────────────────────────────────
# 1. ValueMatchRule
# ──────────────────────────────────────────────

def test_value_match_pass():
    rule = ValueMatchRule("vm", [{"value": "100"}, {"value": "100"}, {"value": "100"}])
    result = rule.check()
    assert result["rule"] == "vm"
    assert result["status"] == "통과"


def test_value_match_fail():
    rule = ValueMatchRule("vm", [{"value": "100"}, {"value": "200"}])
    result = rule.check()
    assert result["rule"] == "vm"
    assert result["status"] == "실패"


# ──────────────────────────────────────────────
# 2. OrderCheckRule
# ──────────────────────────────────────────────

def test_order_pass():
    rule = OrderCheckRule(
        "order",
        [{"value": "2025.01.01"}, {"value": "2025.02.01"}, {"value": "2025.03.01"}],
    )
    result = rule.check()
    assert result["rule"] == "order"
    assert result["status"] == "통과"


def test_order_fail():
    rule = OrderCheckRule(
        "order",
        [{"value": "2025.03.01"}, {"value": "2025.01.01"}],
    )
    result = rule.check()
    assert result["rule"] == "order"
    assert result["status"] == "실패"


# ──────────────────────────────────────────────
# 3. FormulaCheckRule
# ──────────────────────────────────────────────

def test_formula_pass():
    rule = FormulaCheckRule(
        "formula",
        {
            "operands": [{"value": "10"}, {"value": "1500000"}],
            "operator": "*",
            "expected": {"value": "15000000"},
        },
    )
    result = rule.check()
    assert result["rule"] == "formula"
    assert result["status"] == "통과"


def test_formula_warn():
    """계산값 15000000, 기대값 15000001 → 차이 1 → 경고."""
    rule = FormulaCheckRule(
        "formula",
        {
            "operands": [{"value": "10"}, {"value": "1500000"}],
            "operator": "*",
            "expected": {"value": "15000001"},
        },
    )
    result = rule.check()
    assert result["rule"] == "formula"
    assert result["status"] == "경고"


# ──────────────────────────────────────────────
# 4. ExistsRule
# ──────────────────────────────────────────────

def test_exists_pass():
    rule = ExistsRule("exists", [{"value": "abc"}, {"value": "def"}])
    result = rule.check()
    assert result["rule"] == "exists"
    assert result["status"] == "통과"


def test_exists_fail():
    rule = ExistsRule("exists", [{"value": "abc"}, {"value": ""}])
    result = rule.check()
    assert result["rule"] == "exists"
    assert result["status"] == "실패"


# ──────────────────────────────────────────────
# 5. ContainsRule
# ──────────────────────────────────────────────

def test_contains_pass():
    rule = ContainsRule(
        "contains",
        source={"value": "정비"},
        target={"value": "정비비용정산서"},
    )
    result = rule.check()
    assert result["rule"] == "contains"
    assert result["status"] == "통과"


def test_contains_fail():
    rule = ContainsRule(
        "contains",
        source={"value": "수리"},
        target={"value": "정비비용정산서"},
    )
    result = rule.check()
    assert result["rule"] == "contains"
    assert result["status"] == "실패"


# ──────────────────────────────────────────────
# 6. RangeCheckRule
# ──────────────────────────────────────────────

def test_range_check_pass():
    rule = RangeCheckRule(
        "range",
        target={"value": "2025.02.15"},
        bounds={"min": "2025.01.01", "max": "2025.03.31"},
    )
    result = rule.check()
    assert result["rule"] == "range"
    assert result["status"] == "통과"


def test_range_check_fail():
    rule = RangeCheckRule(
        "range",
        target={"value": "2025.05.01"},
        bounds={"min": "2025.01.01", "max": "2025.03.31"},
    )
    result = rule.check()
    assert result["rule"] == "range"
    assert result["status"] == "실패"


# ──────────────────────────────────────────────
# 7. CrossValidator
# ──────────────────────────────────────────────

def test_cross_validator():
    """ValueMatchRule(통과) + ExistsRule(실패) 2개 룰 조합."""
    cv = CrossValidator()
    cv.add_rule(ValueMatchRule("vm", [{"value": "100"}, {"value": "100"}]))
    cv.add_rule(ExistsRule("exists", [{"value": "abc"}, {"value": ""}]))

    results = cv.validate()

    assert len(results) == 2
    assert results[0]["rule"] == "vm"
    assert results[0]["status"] == "통과"
    assert results[1]["rule"] == "exists"
    assert results[1]["status"] == "실패"


def test_cross_validator_process():
    """process()가 context["validation_results"]에 결과를 저장하는지 확인."""
    cv = CrossValidator()
    cv.add_rule(ValueMatchRule("vm", [{"value": "42"}, {"value": "42"}]))
    cv.add_rule(
        FormulaCheckRule(
            "formula",
            {
                "operands": [{"value": "3"}, {"value": "4"}],
                "operator": "+",
                "expected": {"value": "7"},
            },
        )
    )

    context = {"errors": []}
    context = cv.process(doc=None, context=context)

    assert "validation_results" in context
    assert len(context["validation_results"]) == 2
    assert context["validation_results"][0]["status"] == "통과"
    assert context["validation_results"][1]["status"] == "통과"
