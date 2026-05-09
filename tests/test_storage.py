"""
test_storage.py — SQLite Storage CRUD 테스트
TDD 기반으로 storage.py 구현 전 먼저 작성
"""
import os
import tempfile
import pytest

from doc_intelligence.storage import Storage


@pytest.fixture
def storage():
    """임시 파일 기반 Storage 인스턴스 (테스트 격리)"""
    with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
        db_path = f.name
    s = Storage(db_path)
    yield s
    s.close()
    os.unlink(db_path)


# ──────────────────────────────────────────────
# 1. 초기화
# ──────────────────────────────────────────────
def test_init_creates_tables(storage):
    """5개 테이블이 정상 생성되는지 확인"""
    tables = storage.list_tables()
    assert "templates" in tables
    assert "rules" in tables
    assert "presets" in tables
    assert "documents" in tables
    assert "validation_results" in tables


# ──────────────────────────────────────────────
# 2. Templates CRUD
# ──────────────────────────────────────────────
def test_template_crud(storage):
    """템플릿 저장 / 조회 / 수정 / 삭제 전체 흐름"""
    # 저장
    tid = storage.save_template(
        name="정비비용정산서",
        fields=["항목", "금액", "비고"],
        metadata={"version": 1}
    )
    assert tid is not None

    # 조회
    tmpl = storage.get_template(tid)
    assert tmpl["name"] == "정비비용정산서"
    assert tmpl["fields"] == ["항목", "금액", "비고"]
    assert tmpl["metadata"]["version"] == 1
    assert tmpl["match_count"] == 0

    # 수정
    storage.update_template(tid, name="수정된정산서", metadata={"version": 2})
    tmpl_updated = storage.get_template(tid)
    assert tmpl_updated["name"] == "수정된정산서"
    assert tmpl_updated["metadata"]["version"] == 2

    # match_count 증가
    storage.increment_match_count(tid)
    storage.increment_match_count(tid)
    tmpl_counted = storage.get_template(tid)
    assert tmpl_counted["match_count"] == 2

    # 삭제
    storage.delete_template(tid)
    assert storage.get_template(tid) is None


# ──────────────────────────────────────────────
# 3. Rules CRUD
# ──────────────────────────────────────────────
def test_rule_crud(storage):
    """규칙 저장 / 조회 / 수정 / 삭제 전체 흐름"""
    # 저장
    rid = storage.save_rule(
        name="금액 합계 검증",
        rule_type="sum_check",
        conditions={"column": "금액", "operator": "sum_equals", "target": "합계"},
        actions={"on_fail": "경고"}
    )
    assert rid is not None

    # 조회
    rule = storage.get_rule(rid)
    assert rule["name"] == "금액 합계 검증"
    assert rule["rule_type"] == "sum_check"
    assert rule["conditions"]["column"] == "금액"
    assert rule["actions"]["on_fail"] == "경고"

    # 수정
    storage.update_rule(rid, name="수정된 규칙", actions={"on_fail": "실패"})
    rule_updated = storage.get_rule(rid)
    assert rule_updated["name"] == "수정된 규칙"
    assert rule_updated["actions"]["on_fail"] == "실패"

    # 삭제
    storage.delete_rule(rid)
    assert storage.get_rule(rid) is None


# ──────────────────────────────────────────────
# 4. Presets CRUD
# ──────────────────────────────────────────────
def test_preset_crud(storage):
    """프리셋 저장 / 조회 / 수정 (template_ids 포함)"""
    # 먼저 템플릿 2개 저장
    tid1 = storage.save_template(name="템플릿A", fields=["항목"], metadata={})
    tid2 = storage.save_template(name="템플릿B", fields=["금액"], metadata={})

    # 프리셋 저장
    pid = storage.save_preset(
        name="기본 프리셋",
        template_ids=[tid1, tid2],
        rule_ids=[],
        settings={"auto_detect": True}
    )
    assert pid is not None

    # 조회
    preset = storage.get_preset(pid)
    assert preset["name"] == "기본 프리셋"
    assert set(preset["template_ids"]) == {tid1, tid2}
    assert preset["settings"]["auto_detect"] is True

    # 수정
    storage.update_preset(pid, name="수정된 프리셋", settings={"auto_detect": False})
    preset_updated = storage.get_preset(pid)
    assert preset_updated["name"] == "수정된 프리셋"
    assert preset_updated["settings"]["auto_detect"] is False

    # 삭제
    storage.delete_preset(pid)
    assert storage.get_preset(pid) is None


# ──────────────────────────────────────────────
# 5. Validation Results 저장/조회
# ──────────────────────────────────────────────
def test_validation_result(storage):
    """검증 결과 저장 및 조회"""
    # 선행 문서 2개 저장
    doc_id1 = storage.save_document(
        filename="settlement_2024_01.xlsx",
        filepath="/data/settlement_2024_01.xlsx",
        template_id=None,
        parsed_data={"rows": 10}
    )
    doc_id2 = storage.save_document(
        filename="settlement_2024_02.xlsx",
        filepath="/data/settlement_2024_02.xlsx",
        template_id=None,
        parsed_data={"rows": 5}
    )
    assert doc_id1 is not None
    assert doc_id2 is not None

    # 프리셋 저장 (preset_id 필터링 테스트용)
    pid = storage.save_preset(
        name="검증용 프리셋",
        template_ids=[],
        rule_ids=[],
        settings={}
    )

    # 검증 결과 저장 (한국어 status, document_ids 리스트, preset_id 포함)
    vr_id = storage.save_validation_result(
        preset_id=pid,
        rule_id=None,
        document_ids=[doc_id1, doc_id2],
        status="통과",
        detail={"message": "합계 일치"}
    )
    assert vr_id is not None

    # preset_id 필터링 조회
    results = storage.get_validation_results(preset_id=pid)
    assert len(results) == 1
    assert results[0]["status"] == "통과"
    assert results[0]["detail"]["message"] == "합계 일치"
    assert set(results[0]["document_ids"]) == {doc_id1, doc_id2}

    # 실패/경고 status도 저장 가능한지 확인
    storage.save_validation_result(
        preset_id=pid,
        rule_id=None,
        document_ids=[doc_id1],
        status="실패",
        detail={"message": "금액 불일치"}
    )
    storage.save_validation_result(
        preset_id=pid,
        rule_id=None,
        document_ids=[doc_id2],
        status="경고",
        detail={"message": "빈 셀 감지"}
    )
    all_results = storage.get_validation_results(preset_id=pid)
    assert len(all_results) == 3
    statuses = {r["status"] for r in all_results}
    assert statuses == {"통과", "실패", "경고"}

    # preset_id=None → 전체 조회
    all_global = storage.get_validation_results()
    assert len(all_global) == 3


# ──────────────────────────────────────────────
# 6. 없는 ID 조회 → None 반환
# ──────────────────────────────────────────────
def test_get_nonexistent_returns_none(storage):
    """존재하지 않는 ID 조회 시 None 반환"""
    assert storage.get_template(9999) is None
    assert storage.get_rule(9999) is None
    assert storage.get_preset(9999) is None


# ──────────────────────────────────────────────
# 7. 복수 템플릿 전체 조회
# ──────────────────────────────────────────────
def test_get_all_templates(storage):
    """복수 템플릿 저장 후 전체 목록 조회"""
    storage.save_template(name="템플릿1", fields=["A"], metadata={})
    storage.save_template(name="템플릿2", fields=["B"], metadata={})
    storage.save_template(name="템플릿3", fields=["C"], metadata={})

    all_templates = storage.get_all_templates()
    assert len(all_templates) == 3
    names = {t["name"] for t in all_templates}
    assert names == {"템플릿1", "템플릿2", "템플릿3"}


# ──────────────────────────────────────────────
# 8. find_presets_by_template_ids — 프리셋 자동 감지
# ──────────────────────────────────────────────
def test_find_preset_by_template_ids(storage):
    """
    열린 문서들의 template_id 조합이 프리셋의 template_ids의 부분집합인지 확인.
    예: 프리셋이 [A, B, C]를 요구할 때
        - 열린 문서 = [A, B, C] → 매칭 (정확히 일치)
        - 열린 문서 = [A, B]    → 매칭 (부분집합)
        - 열린 문서 = [A, D]    → 미매칭 (D가 프리셋에 없음)
    """
    tid_a = storage.save_template(name="A", fields=[], metadata={})
    tid_b = storage.save_template(name="B", fields=[], metadata={})
    tid_c = storage.save_template(name="C", fields=[], metadata={})
    tid_d = storage.save_template(name="D", fields=[], metadata={})

    pid = storage.save_preset(
        name="ABC 프리셋",
        template_ids=[tid_a, tid_b, tid_c],
        rule_ids=[],
        settings={}
    )

    # 정확히 일치 → 매칭
    matched = storage.find_presets_by_template_ids([tid_a, tid_b, tid_c])
    assert any(p["id"] == pid for p in matched)

    # 부분집합 → 매칭
    matched_partial = storage.find_presets_by_template_ids([tid_a, tid_b])
    assert any(p["id"] == pid for p in matched_partial)

    # D가 포함 → 미매칭 (열린 문서 중 프리셋에 없는 항목이 있으면 제외)
    not_matched = storage.find_presets_by_template_ids([tid_a, tid_d])
    assert not any(p["id"] == pid for p in not_matched)
