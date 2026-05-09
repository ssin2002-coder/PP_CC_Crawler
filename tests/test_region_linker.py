"""
test_region_linker.py — RegionLinker TDD
4개 테스트:
  1. test_linked_region         : LinkedRegion 생성
  2. test_create_rule           : 2개 영역으로 룰 생성, name/rule_type/regions 확인
  3. test_dpi_scaling           : 150% DPI에서 300 → 200
  4. test_identify_app_from_hwnd: COM 없는 환경에서 None 반환
"""
import pytest

from doc_intelligence.region_linker import LinkedRegion, RegionLinker


# ──────────────────────────────────────────────
# Fixtures
# ──────────────────────────────────────────────

@pytest.fixture
def linker():
    """storage 없이 RegionLinker 인스턴스 생성"""
    return RegionLinker(storage=None)


def _make_region(app="EXCEL.EXE", doc="test.xlsx", loc="Sheet1!B4", rect=(0, 0, 100, 50)):
    """테스트용 LinkedRegion 헬퍼"""
    return LinkedRegion(
        app_name=app,
        doc_name=doc,
        location=loc,
        screen_rect=rect,
        screenshot=None,
    )


# ──────────────────────────────────────────────
# 1. LinkedRegion 생성
# ──────────────────────────────────────────────

def test_linked_region():
    """LinkedRegion 데이터클래스 필드 정상 생성 확인"""
    region = LinkedRegion(
        app_name="EXCEL.EXE",
        doc_name="정비비용정산.xlsx",
        location="Sheet1!B4",
        screen_rect=(100, 200, 300, 50),
        screenshot=None,
    )
    assert region.app_name == "EXCEL.EXE"
    assert region.doc_name == "정비비용정산.xlsx"
    assert region.location == "Sheet1!B4"
    assert region.screen_rect == (100, 200, 300, 50)
    assert region.screenshot is None


# ──────────────────────────────────────────────
# 2. create_rule — 2개 영역으로 룰 생성
# ──────────────────────────────────────────────

def test_create_rule(linker):
    """2개 LinkedRegion으로 룰 생성 후 name/rule_type/regions 필드 확인"""
    r1 = _make_region(loc="Sheet1!B4", rect=(0, 0, 100, 50))
    r2 = _make_region(app="WINWORD.EXE", doc="계약서.docx", loc="para:3", rect=(200, 100, 150, 60))

    rule = linker.create_rule([r1, r2], rule_type="cross_ref", rule_name="엑셀-워드 교차검증")

    assert rule["name"] == "엑셀-워드 교차검증"
    assert rule["rule_type"] == "cross_ref"
    assert len(rule["regions"]) == 2
    assert rule["regions"][0]["location"] == "Sheet1!B4"
    assert rule["regions"][1]["location"] == "para:3"


# ──────────────────────────────────────────────
# 3. DPI 스케일링 — 150% DPI에서 300 → 200
# ──────────────────────────────────────────────

def test_dpi_scaling(linker):
    """scale_factor=1.5 적용 시 물리 좌표 300 → 논리 좌표 200 반환"""
    result = linker._apply_dpi_scale(300, 1.5)
    assert result == 200


# ──────────────────────────────────────────────
# 4. 앱 식별 — COM 없는 환경에서 None 반환
# ──────────────────────────────────────────────

def test_identify_app_from_hwnd(linker):
    """win32gui/psutil 미설치(또는 유효하지 않은 좌표) 환경에서 None 반환 확인"""
    # 화면 바깥 좌표(-9999, -9999)를 전달하여 앱 식별 불가 상황 재현
    result = linker._get_app_from_point(-9999, -9999)
    assert result is None
