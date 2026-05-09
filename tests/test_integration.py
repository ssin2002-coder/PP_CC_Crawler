"""
test_integration.py — 전체 파이프라인 통합 테스트
5개 테스트:
  1. test_full_pipeline       — Engine + Fingerprinter + EntityExtractor 등록 후 파이프라인 실행
  2. test_learn_and_match     — Fingerprinter 학습 → 동일 문서 매칭 → score >= 0.85
  3. test_cross_validation    — CrossValidator + ValueMatchRule + OrderCheckRule → 모두 통과
  4. test_graph_integration   — DocGraph 2개 문서 + 관계 → to_html 길이 > 0
  5. test_preset_auto_detect  — Storage 템플릿 2개 + 프리셋 1개 → find_presets_by_template_ids 매칭
"""
import os
import tempfile

import pytest

from doc_intelligence.engine import Engine, ParsedDocument, CellData
from doc_intelligence.fingerprint import Fingerprinter
from doc_intelligence.extractor import EntityExtractor
from doc_intelligence.validator import CrossValidator, ValueMatchRule, OrderCheckRule
from doc_intelligence.graph import DocGraph
from doc_intelligence.storage import Storage


# ──────────────────────────────────────────────
# 공통 헬퍼
# ──────────────────────────────────────────────

def _make_doc(file_path: str, labels: list, amount: str = "", date: str = "") -> ParsedDocument:
    """테스트용 ParsedDocument 생성 헬퍼."""
    cells = []
    for i, label in enumerate(labels):
        cells.append(CellData(
            address=f"Sheet1!R{i + 1}C1",
            value=label,
            data_type="text",
            neighbors={},
        ))

    raw_parts = list(labels)

    if amount:
        cells.append(CellData(
            address=f"Sheet1!R{len(labels) + 1}C2",
            value=amount,
            data_type="text",
            neighbors={},
        ))
        raw_parts.append(amount)

    if date:
        cells.append(CellData(
            address=f"Sheet1!R{len(labels) + 2}C2",
            value=date,
            data_type="text",
            neighbors={},
        ))
        raw_parts.append(date)

    return ParsedDocument(
        file_path=file_path,
        file_type="excel",
        raw_text=" ".join(raw_parts),
        structure={"merge_cells": []},
        cells=cells,
        metadata={},
    )


@pytest.fixture
def db_path():
    """임시 DB 경로 — 테스트 격리."""
    with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
        path = f.name
    yield path
    try:
        os.unlink(path)
    except OSError:
        pass


# ──────────────────────────────────────────────
# 1. test_full_pipeline
# ──────────────────────────────────────────────

def test_full_pipeline(db_path):
    """Engine + Fingerprinter + EntityExtractor 등록 후 파이프라인 실행.
    context에 fingerprint, entities 키가 존재해야 한다.
    """
    engine = Engine(db_path=db_path)
    engine.register(Fingerprinter())
    engine.register(EntityExtractor())

    doc = _make_doc(
        file_path="test_full.xlsx",
        labels=["항목", "금액", "업체명", "날짜"],
        amount="1,500,000원",
        date="2025.03.15",
    )

    context = engine.process(doc)

    assert "fingerprint" in context, "context에 fingerprint 키가 없음"
    assert "entities" in context, "context에 entities 키가 없음"
    assert isinstance(context["fingerprint"], dict)
    assert isinstance(context["entities"], list)
    # 엔티티 최소 1개 이상 추출 (금액, 날짜 포함)
    assert len(context["entities"]) >= 1


# ──────────────────────────────────────────────
# 2. test_learn_and_match
# ──────────────────────────────────────────────

def test_learn_and_match(db_path):
    """Fingerprinter에 문서 학습 → 같은 문서 매칭 → score >= 0.85."""
    engine = Engine(db_path=db_path)
    fp = Fingerprinter()
    engine.register(fp)

    doc = _make_doc(
        file_path="learn_doc.xlsx",
        labels=["항목", "금액", "비고", "합계", "세금", "총액", "담당자", "승인"],
    )

    fp.learn(doc, "정비비용정산서")

    result = fp.match(doc)

    assert result["template"] is not None, "template_id가 None"
    assert result["score"] >= 0.85, f"score {result['score']:.3f} < 0.85"
    assert result["auto"] is True, "auto 플래그가 False"


# ──────────────────────────────────────────────
# 3. test_cross_validation
# ──────────────────────────────────────────────

def test_cross_validation():
    """CrossValidator + ValueMatchRule + OrderCheckRule → 모두 통과."""
    validator = CrossValidator()

    # ValueMatchRule: 두 값이 동일 → 통과
    value_rule = ValueMatchRule(
        name="금액일치",
        regions=[{"value": "1500000"}, {"value": "1500000"}],
    )

    # OrderCheckRule: 날짜 오름차순 → 통과
    order_rule = OrderCheckRule(
        name="날짜순서",
        regions=[
            {"value": "2025.01.15"},
            {"value": "2025.02.20"},
            {"value": "2025.03.30"},
        ],
    )

    validator.add_rule(value_rule)
    validator.add_rule(order_rule)

    results = validator.validate()

    assert len(results) == 2, f"결과 개수가 2가 아님: {len(results)}"
    statuses = [r["status"] for r in results]
    assert all(s == "통과" for s in statuses), f"통과가 아닌 결과 존재: {statuses}"


# ──────────────────────────────────────────────
# 4. test_graph_integration
# ──────────────────────────────────────────────

def test_graph_integration():
    """DocGraph에 2개 문서 + 관계 추가 → to_html 길이 > 0."""
    graph = DocGraph()

    graph.add_document("문서A.xlsx", entities=["금액: 1,500,000원", "날짜: 2025.03.15"])
    graph.add_document("문서B.xlsx", entities=["금액: 1,500,000원", "업체명: 삼성전자"])

    graph.add_relationship("문서A.xlsx", "문서B.xlsx", rule_name="금액일치", status="pass")

    assert graph.node_count() == 2, f"노드 수가 2가 아님: {graph.node_count()}"

    edges = graph.get_edges()
    assert len(edges) == 1, f"엣지 수가 1이 아님: {len(edges)}"
    assert edges[0][2]["status"] == "pass"

    html = graph.to_html()
    assert len(html) > 0, "to_html() 결과가 빈 문자열"
    assert isinstance(html, str), "to_html() 반환값이 str이 아님"


# ──────────────────────────────────────────────
# 5. test_preset_auto_detect
# ──────────────────────────────────────────────

def test_preset_auto_detect(db_path):
    """Storage에 템플릿 2개 + 프리셋 1개 → find_presets_by_template_ids 매칭."""
    storage = Storage(db_path=db_path)

    # 템플릿 2개 등록
    tid1 = storage.save_template(
        name="정비비용정산서",
        fields=["항목", "금액", "합계"],
        metadata={"label_positions": {}, "merge_pattern": ""},
    )
    tid2 = storage.save_template(
        name="검수확인서",
        fields=["검수자", "검수일", "설비코드"],
        metadata={"label_positions": {}, "merge_pattern": ""},
    )

    # 두 템플릿을 포함하는 프리셋 1개 등록
    preset_id = storage.save_preset(
        name="정비비용_검수_세트",
        template_ids=[tid1, tid2],
        rule_ids=[],
        settings={},
    )

    # [tid1, tid2] 로 find_presets_by_template_ids → 매칭되어야 함
    matched = storage.find_presets_by_template_ids([tid1, tid2])
    assert len(matched) >= 1, "프리셋이 매칭되지 않음"
    assert any(p["id"] == preset_id for p in matched), "등록한 프리셋이 결과에 없음"

    # tid1 만으로도 subset이므로 매칭되어야 함
    matched_partial = storage.find_presets_by_template_ids([tid1])
    assert len(matched_partial) >= 1, "부분집합 매칭 실패"

    # tid1, tid2 외 존재하지 않는 tid로 조회 시 매칭 없어야 함
    matched_none = storage.find_presets_by_template_ids([tid1, tid2, 9999])
    assert len(matched_none) == 0, "존재하지 않는 template_id가 포함될 때 매칭되면 안 됨"

    storage.close()
