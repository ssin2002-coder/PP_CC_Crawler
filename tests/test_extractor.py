"""
test_extractor.py — EntityExtractor regex + 인접 셀 기반 엔티티 추출기 TDD
TDD 기반으로 extractor.py 구현 전 먼저 작성
"""
import pytest

from doc_intelligence.engine import CellData, ParsedDocument
from doc_intelligence.extractor import EntityExtractor


# ──────────────────────────────────────────────
# Fixtures
# ──────────────────────────────────────────────

def _make_doc(cell_values):
    """테스트용 ParsedDocument 생성 헬퍼 (셀 11개 기준 문서)"""
    cells = [
        CellData(
            address=f"Sheet1!R{i+1}C1",
            value=v,
            data_type="text",
            neighbors={},
        )
        for i, v in enumerate(cell_values)
    ]
    return ParsedDocument(
        file_path="test.xlsx",
        file_type="excel",
        raw_text=" ".join(str(v) for v in cell_values),
        structure={},
        cells=cells,
        metadata={},
    )


# 표준 테스트 문서: 11개 셀
#  견적서, 업체명, 삼성엔지니어링, 합계, 15,000,000원, 날짜, 2025.03.15, 설비코드, PP-2045, 사업자번호, 123-45-67890
STANDARD_CELLS = [
    "견적서",
    "업체명",
    "삼성엔지니어링",
    "합계",
    "15,000,000원",
    "날짜",
    "2025.03.15",
    "설비코드",
    "PP-2045",
    "사업자번호",
    "123-45-67890",
]


@pytest.fixture
def extractor():
    """EntityExtractor 인스턴스"""
    e = EntityExtractor()
    e.initialize(None)
    return e


@pytest.fixture
def standard_doc():
    """표준 테스트 문서"""
    return _make_doc(STANDARD_CELLS)


# ──────────────────────────────────────────────
# 1. 금액 추출
# ──────────────────────────────────────────────

def test_extract_amount(extractor, standard_doc):
    """'15,000,000원' 금액 엔티티 추출 확인"""
    entities = extractor.extract(standard_doc)
    amounts = [e for e in entities if e.type == "금액"]

    assert len(amounts) >= 1
    assert any("15,000,000원" in e.value for e in amounts)


# ──────────────────────────────────────────────
# 2. 날짜 추출
# ──────────────────────────────────────────────

def test_extract_date(extractor, standard_doc):
    """'2025.03.15' 날짜 엔티티 추출 확인"""
    entities = extractor.extract(standard_doc)
    dates = [e for e in entities if e.type == "날짜"]

    assert len(dates) >= 1
    assert any("2025.03.15" in e.value for e in dates)


# ──────────────────────────────────────────────
# 3. 설비코드 추출
# ──────────────────────────────────────────────

def test_extract_equipment_code(extractor, standard_doc):
    """'PP-2045' 설비코드 엔티티 추출 확인"""
    entities = extractor.extract(standard_doc)
    codes = [e for e in entities if e.type == "설비코드"]

    assert len(codes) >= 1
    assert any("PP-2045" in e.value for e in codes)


# ──────────────────────────────────────────────
# 4. 사업자번호 추출
# ──────────────────────────────────────────────

def test_extract_biz_number(extractor, standard_doc):
    """'123-45-67890' 사업자번호 엔티티 추출 확인"""
    entities = extractor.extract(standard_doc)
    biz = [e for e in entities if e.type == "사업자번호"]

    assert len(biz) >= 1
    assert any("123-45-67890" in e.value for e in biz)


# ──────────────────────────────────────────────
# 5. 인접 셀 힌트 기반 confidence 상향
# ──────────────────────────────────────────────

def test_neighbor_inference(extractor, standard_doc):
    """'합계' 레이블 다음 셀 '15,000,000원'의 confidence >= 0.8 확인"""
    entities = extractor.extract(standard_doc)
    amounts = [e for e in entities if e.type == "금액" and "15,000,000원" in e.value]

    assert len(amounts) >= 1
    # '합계' 인접 셀이므로 confidence가 boost되어 0.8 이상이어야 함
    assert amounts[0].confidence >= 0.8
