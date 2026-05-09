"""
test_ui.py — ui_components.py Mock 기반 테스트 (tkinter 없이 동작)
테스트 3개:
  1. LearningModeDialog 생성 / get_corrected_mappings
  2. RuleManagerWidget 생성
  3. ValidationResultWidget 생성
"""
from unittest.mock import MagicMock, patch


# ──────────────────────────────────────────────
# 공통 픽스처 헬퍼
# ──────────────────────────────────────────────

def _sample_entities():
    return [
        {"value": "2025.01.01", "type": "날짜"},
        {"value": "삼성전자", "type": "업체명"},
        {"value": "1,500,000원", "type": "금액"},
    ]


def _sample_presets():
    return [
        {"id": 1, "name": "정비비용_프리셋"},
        {"id": 2, "name": "검수_프리셋"},
    ]


def _sample_rules():
    return [
        {"id": 1, "name": "금액일치", "rule_type": "ValueMatchRule"},
        {"id": 2, "name": "날짜순서", "rule_type": "OrderCheckRule"},
    ]


def _sample_results():
    return [
        {"rule": "금액일치", "status": "통과", "detail": "모든 값 일치: '1500000'"},
        {"rule": "날짜순서", "status": "실패", "detail": "날짜 역순 감지"},
    ]


# ──────────────────────────────────────────────
# 테스트 1: LearningModeDialog 생성 / 매핑
# ──────────────────────────────────────────────

def test_learning_dialog_create_and_mapping():
    """parent=None 헤드리스 환경에서 LearningModeDialog 생성 및 get_corrected_mappings 확인."""
    from doc_intelligence.ui_components import LearningModeDialog

    entities = _sample_entities()
    dialog = LearningModeDialog(parent=None, entities=entities, doc_name="test.xlsx")

    # 인스턴스 생성 확인
    assert dialog is not None
    assert dialog.doc_name == "test.xlsx"
    assert dialog.entities is entities

    # get_corrected_mappings: 보정 없을 때 기존 type 반환
    mappings = dialog.get_corrected_mappings()
    assert isinstance(mappings, dict)
    assert len(mappings) == 3
    assert mappings[0] == "날짜"
    assert mappings[1] == "업체명"
    assert mappings[2] == "금액"


def test_learning_dialog_with_corrections():
    """_corrections 직접 주입 시 get_corrected_mappings가 보정값을 반환하는지 확인."""
    from doc_intelligence.ui_components import LearningModeDialog

    entities = _sample_entities()
    dialog = LearningModeDialog(parent=None, entities=entities)

    # 보정 직접 주입
    dialog._corrections = {0: "착공일", 1: "업체명", 2: "예상비용"}
    mappings = dialog.get_corrected_mappings()

    assert mappings[0] == "착공일"
    assert mappings[1] == "업체명"
    assert mappings[2] == "예상비용"


# ──────────────────────────────────────────────
# 테스트 2: RuleManagerWidget 생성
# ──────────────────────────────────────────────

def test_rule_manager_create():
    """parent=None 헤드리스 환경에서 RuleManagerWidget 생성 확인."""
    from doc_intelligence.ui_components import RuleManagerWidget

    presets = _sample_presets()
    rules = _sample_rules()

    widget = RuleManagerWidget(parent=None, presets=presets, rules=rules)

    assert widget is not None
    assert widget.presets is presets
    assert widget.rules is rules
    assert widget.get_selected_preset() is None
    assert widget.get_selected_rule() is None


# ──────────────────────────────────────────────
# 테스트 3: ValidationResultWidget 생성
# ──────────────────────────────────────────────

def test_validation_result_create():
    """parent=None 헤드리스 환경에서 ValidationResultWidget 생성 확인."""
    from doc_intelligence.ui_components import ValidationResultWidget

    results = _sample_results()
    widget = ValidationResultWidget(parent=None, results=results)

    assert widget is not None
    assert widget.results is results
    assert len(widget.results) == 2
    assert widget.results[0]["status"] == "통과"
    assert widget.results[1]["status"] == "실패"
