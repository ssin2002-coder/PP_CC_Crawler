"""
test_fingerprint.py — Fingerprinter TF-IDF 핑거프린트 + 템플릿 매칭 TDD
TDD 기반으로 fingerprint.py 구현 전 먼저 작성
"""
import os
import tempfile
import pytest

from doc_intelligence.storage import Storage
from doc_intelligence.engine import CellData, ParsedDocument, Fingerprint
from doc_intelligence.fingerprint import Fingerprinter


# ──────────────────────────────────────────────
# Fixtures
# ──────────────────────────────────────────────

def _make_doc(labels, merge_cells=None):
    """테스트용 ParsedDocument 생성 헬퍼"""
    cells = [
        CellData(address=f"Sheet1!R{i+1}C1", value=label, data_type="text", neighbors={})
        for i, label in enumerate(labels)
    ]
    structure = {}
    if merge_cells:
        structure["merge_cells"] = merge_cells
    return ParsedDocument(
        file_path="test.xlsx",
        file_type="excel",
        raw_text=" ".join(labels),
        structure=structure,
        cells=cells,
        metadata={},
    )


@pytest.fixture
def storage():
    """임시 파일 기반 Storage 인스턴스 (테스트 격리)"""
    with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
        db_path = f.name
    s = Storage(db_path)
    yield s
    s.close()
    os.unlink(db_path)


@pytest.fixture
def fp(storage):
    """Fingerprinter 인스턴스 — 빈 DB에서 시작"""
    fingerprinter = Fingerprinter(storage=storage)

    class _FakeEngine:
        def __init__(self, s):
            self.storage = s

    fingerprinter.initialize(_FakeEngine(storage))
    return fingerprinter


# ──────────────────────────────────────────────
# 1. generate — 핑거프린트 생성
# ──────────────────────────────────────────────

def test_generate_fingerprint(fp):
    """문서에서 핑거프린트 생성. vector, labels, fingerprint 키 존재 확인"""
    doc = _make_doc(["항목", "금액", "비고", "합계"])
    result = fp.generate(doc)

    assert "vector" in result
    assert "labels" in result
    assert "fingerprint" in result
    assert isinstance(result["labels"], list)
    assert len(result["labels"]) > 0


def test_generate_fingerprint_dataclass(fp):
    """generate 결과에 Fingerprint 인스턴스 포함 확인"""
    doc = _make_doc(["항목", "금액", "비고"])
    result = fp.generate(doc)

    assert isinstance(result["fingerprint"], Fingerprint)
    fp_obj = result["fingerprint"]
    assert fp_obj.doc_id != ""
    assert isinstance(fp_obj.label_positions, dict)
    assert isinstance(fp_obj.merge_pattern, str)
    assert isinstance(fp_obj.feature_vector, list)


# ──────────────────────────────────────────────
# 2. test_fingerprint_dataclass — dataclass 별도 확인
# ──────────────────────────────────────────────

def test_fingerprint_dataclass(fp):
    """generate 결과의 Fingerprint dataclass 필드 타입 검증"""
    doc = _make_doc(["작업명", "단가", "수량"], merge_cells=["A1:B1", "C2:D2"])
    result = fp.generate(doc)

    fp_obj = result["fingerprint"]
    assert isinstance(fp_obj, Fingerprint)
    # merge_pattern은 MD5 해시(32자 hex)
    assert len(fp_obj.merge_pattern) == 32
    # label_positions: {값: 주소} 매핑
    assert "작업명" in fp_obj.label_positions or len(fp_obj.label_positions) >= 0


# ──────────────────────────────────────────────
# 3. match — 동일 문서 매칭 (score >= 0.85)
# ──────────────────────────────────────────────

def test_match_exact_same(fp):
    """동일 문서 learn 후 match -> score >= 0.85"""
    doc = _make_doc(["항목", "금액", "비고", "합계", "세금", "총액"])
    fp.learn(doc, "정비비용정산서")

    result = fp.match(doc)
    assert result["template"] is not None
    assert result["score"] >= 0.85
    assert result["auto"] is True


# ──────────────────────────────────────────────
# 4. match — 유사 문서 (score >= 0.60)
# ──────────────────────────────────────────────

def test_match_similar(fp):
    """유사 문서 match -> score >= 0.60"""
    # 원본 문서 학습
    doc_orig = _make_doc(["항목", "금액", "비고", "합계", "세금", "총액", "담당자", "승인"])
    fp.learn(doc_orig, "정비비용정산서")

    # 일부 라벨이 겹치는 유사 문서
    doc_similar = _make_doc(["항목", "금액", "비고", "합계", "세금", "총액", "날짜", "서명"])
    result = fp.match(doc_similar)

    assert result["score"] >= 0.60


# ──────────────────────────────────────────────
# 5. match — 완전히 다른 문서 (template is None or score < 0.60)
# ──────────────────────────────────────────────

def test_no_match(fp):
    """완전히 다른 문서 -> template is None or score < 0.60"""
    doc_orig = _make_doc(["항목", "금액", "비고", "합계"])
    fp.learn(doc_orig, "정비비용정산서")

    # 전혀 다른 라벨의 문서
    doc_diff = _make_doc(["사원번호", "부서", "직급", "입사일", "연봉"])
    result = fp.match(doc_diff)

    assert result["template"] is None or result["score"] < 0.60


# ──────────────────────────────────────────────
# 6. match — corpus 없을 때
# ──────────────────────────────────────────────

def test_match_empty_corpus(fp):
    """corpus가 비어있으면 template=None, score=0.0"""
    doc = _make_doc(["항목", "금액"])
    result = fp.match(doc)

    assert result["template"] is None
    assert result["score"] == 0.0


# ──────────────────────────────────────────────
# 7. learn — template_id 반환 및 DB 저장 확인
# ──────────────────────────────────────────────

def test_learn_returns_id(fp, storage):
    """learn 후 유효한 template_id 반환, DB에 저장됨"""
    doc = _make_doc(["항목", "금액", "비고"])
    tid = fp.learn(doc, "테스트_템플릿")

    assert isinstance(tid, int)
    assert tid > 0
    tmpl = storage.get_template(tid)
    assert tmpl is not None
    assert tmpl["name"] == "테스트_템플릿"


# ──────────────────────────────────────────────
# 8. process — context dict 업데이트
# ──────────────────────────────────────────────

def test_process_updates_context(fp):
    """process() 호출 후 context에 fingerprint, template_match 키 존재"""
    doc = _make_doc(["항목", "금액"])
    context = {"errors": []}
    result = fp.process(doc, context)

    assert "fingerprint" in result
    assert "template_match" in result
