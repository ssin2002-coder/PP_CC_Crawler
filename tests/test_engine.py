"""
test_engine.py — Engine 코어 파이프라인 + 데이터 모델 테스트
TDD: 구현 전에 작성된 테스트 파일
"""
import pytest
from doc_intelligence.engine import (
    CellData,
    ParsedDocument,
    Entity,
    Fingerprint,
    Engine,
)


# ──────────────────────────────────────────────
# 1. 데이터클래스 생성 테스트
# ──────────────────────────────────────────────
class TestDataclasses:
    def test_cell_data_creation(self):
        cell = CellData(
            address="Sheet1!R2C3",
            value=12345,
            data_type="number",
            neighbors={"left": "Sheet1!R2C2", "right": "Sheet1!R2C4"},
        )
        assert cell.address == "Sheet1!R2C3"
        assert cell.value == 12345
        assert cell.data_type == "number"
        assert cell.neighbors["left"] == "Sheet1!R2C2"

    def test_parsed_document_creation(self):
        doc = ParsedDocument(
            file_path="/tmp/sample.xlsx",
            file_type="excel",
            raw_text="원자재비 1,000,000",
            structure={"sheets": ["Sheet1"]},
            cells=[],
            metadata={"author": "테스트"},
        )
        assert doc.file_path == "/tmp/sample.xlsx"
        assert doc.file_type == "excel"
        assert doc.raw_text == "원자재비 1,000,000"
        assert doc.structure == {"sheets": ["Sheet1"]}
        assert doc.cells == []
        assert doc.metadata["author"] == "테스트"

    def test_entity_creation(self):
        entity = Entity(
            type="금액",
            value="1,000,000",
            location="Sheet1!R3C5",
            confidence=0.95,
        )
        assert entity.type == "금액"
        assert entity.value == "1,000,000"
        assert entity.location == "Sheet1!R3C5"
        assert entity.confidence == 0.95

    def test_fingerprint_creation(self):
        fp = Fingerprint(
            doc_id="abc-123",
            feature_vector=[0.1, 0.2, 0.3],
            label_positions={"업체명": "R1C1", "금액": "R3C5"},
            merge_pattern="A1:C1",
        )
        assert fp.doc_id == "abc-123"
        assert fp.feature_vector == [0.1, 0.2, 0.3]
        assert fp.label_positions["업체명"] == "R1C1"
        assert fp.merge_pattern == "A1:C1"

    def test_cell_data_various_types(self):
        """CellData의 value는 str, int, float 등 다양한 타입 허용"""
        str_cell = CellData(address="para:1", value="텍스트", data_type="text", neighbors={})
        float_cell = CellData(address="slide1:shape1", value=3.14, data_type="number", neighbors={})
        assert str_cell.value == "텍스트"
        assert float_cell.value == 3.14

    def test_entity_confidence_range(self):
        """confidence는 0.0 ~ 1.0 범위"""
        low = Entity(type="날짜", value="2024-01-01", location="R1C1", confidence=0.0)
        high = Entity(type="업체명", value="삼성전자", location="R2C1", confidence=1.0)
        assert low.confidence == 0.0
        assert high.confidence == 1.0


# ──────────────────────────────────────────────
# 2. 플러그인 등록 / 목록 테스트
# ──────────────────────────────────────────────
class FakePlugin:
    """테스트용 더미 플러그인"""
    def __init__(self, name: str):
        self.name = name
        self.initialized = False

    def initialize(self, engine):
        self.initialized = True

    def process(self, doc, context: dict) -> dict:
        context[self.name] = "done"
        return context


class TestRegisterAndList:
    def setup_method(self):
        self.engine = Engine(db_path=":memory:")

    def test_register_single_plugin(self):
        plugin = FakePlugin("parser")
        self.engine.register(plugin)
        assert "parser" in self.engine.list_plugins()

    def test_register_multiple_plugins(self):
        p1 = FakePlugin("parser")
        p2 = FakePlugin("extractor")
        p3 = FakePlugin("validator")
        self.engine.register(p1)
        self.engine.register(p2)
        self.engine.register(p3)
        names = self.engine.list_plugins()
        assert names == ["parser", "extractor", "validator"]

    def test_initialize_called_on_register(self):
        plugin = FakePlugin("parser")
        assert plugin.initialized is False
        self.engine.register(plugin)
        assert plugin.initialized is True

    def test_list_plugins_empty(self):
        assert self.engine.list_plugins() == []


# ──────────────────────────────────────────────
# 3. 비활성화 / 활성화 토글 테스트
# ──────────────────────────────────────────────
class TestDisableEnable:
    def setup_method(self):
        self.engine = Engine(db_path=":memory:")
        self.plugin = FakePlugin("parser")
        self.engine.register(self.plugin)

    def test_disable_plugin(self):
        self.engine.disable("parser")
        doc = ParsedDocument(
            file_path="/tmp/test.xlsx",
            file_type="excel",
            raw_text="",
            structure={},
            cells=[],
            metadata={},
        )
        result = self.engine.process(doc)
        # 비활성화된 플러그인은 실행되지 않으므로 context에 키 없음
        assert "parser" not in result

    def test_enable_plugin_after_disable(self):
        self.engine.disable("parser")
        self.engine.enable("parser")
        doc = ParsedDocument(
            file_path="/tmp/test.xlsx",
            file_type="excel",
            raw_text="",
            structure={},
            cells=[],
            metadata={},
        )
        result = self.engine.process(doc)
        assert result.get("parser") == "done"

    def test_disable_nonexistent_plugin_no_error(self):
        """존재하지 않는 플러그인 비활성화 시 오류 없어야 함"""
        self.engine.disable("nonexistent")  # 예외 없어야 함

    def test_enable_nonexistent_plugin_no_error(self):
        """존재하지 않는 플러그인 활성화 시 오류 없어야 함"""
        self.engine.enable("nonexistent")  # 예외 없어야 함


# ──────────────────────────────────────────────
# 4. 파이프라인 실행 순서 테스트
# ──────────────────────────────────────────────
class OrderTrackingPlugin:
    """실행 순서를 기록하는 플러그인"""
    def __init__(self, name: str, order_log: list):
        self.name = name
        self._order_log = order_log

    def initialize(self, engine):
        pass

    def process(self, doc, context: dict) -> dict:
        self._order_log.append(self.name)
        context[self.name] = len(self._order_log)
        return context


class TestPipelineOrder:
    def setup_method(self):
        self.engine = Engine(db_path=":memory:")
        self.order_log = []

    def test_plugins_execute_in_registration_order(self):
        p1 = OrderTrackingPlugin("first", self.order_log)
        p2 = OrderTrackingPlugin("second", self.order_log)
        p3 = OrderTrackingPlugin("third", self.order_log)
        self.engine.register(p1)
        self.engine.register(p2)
        self.engine.register(p3)

        doc = ParsedDocument(
            file_path="/tmp/test.xlsx",
            file_type="excel",
            raw_text="",
            structure={},
            cells=[],
            metadata={},
        )
        result = self.engine.process(doc)

        assert self.order_log == ["first", "second", "third"]
        assert result["first"] == 1
        assert result["second"] == 2
        assert result["third"] == 3

    def test_context_accumulates_across_plugins(self):
        """각 플러그인의 결과가 context에 누적되어야 함"""
        p1 = FakePlugin("step1")
        p2 = FakePlugin("step2")
        self.engine.register(p1)
        self.engine.register(p2)

        doc = ParsedDocument(
            file_path="/tmp/test.xlsx",
            file_type="excel",
            raw_text="",
            structure={},
            cells=[],
            metadata={},
        )
        result = self.engine.process(doc)
        assert "step1" in result
        assert "step2" in result


# ──────────────────────────────────────────────
# 5. 플러그인 예외 처리 테스트
# ──────────────────────────────────────────────
class ExceptionPlugin:
    """process()에서 예외를 발생시키는 플러그인"""
    def __init__(self, name: str):
        self.name = name

    def initialize(self, engine):
        pass

    def process(self, doc, context: dict) -> dict:
        raise RuntimeError(f"{self.name}: 의도적 예외 발생")


class TestPluginExceptionHandling:
    def setup_method(self):
        self.engine = Engine(db_path=":memory:")

    def test_exception_does_not_crash_pipeline(self):
        """예외가 발생해도 파이프라인이 중단되지 않아야 함"""
        bad_plugin = ExceptionPlugin("crasher")
        good_plugin = FakePlugin("survivor")
        self.engine.register(bad_plugin)
        self.engine.register(good_plugin)

        doc = ParsedDocument(
            file_path="/tmp/test.xlsx",
            file_type="excel",
            raw_text="",
            structure={},
            cells=[],
            metadata={},
        )
        # 예외가 전파되지 않아야 함
        result = self.engine.process(doc)
        assert result.get("survivor") == "done"

    def test_exception_recorded_in_errors(self):
        """예외 발생 시 context['errors']에 기록되어야 함"""
        bad_plugin = ExceptionPlugin("crasher")
        self.engine.register(bad_plugin)

        doc = ParsedDocument(
            file_path="/tmp/test.xlsx",
            file_type="excel",
            raw_text="",
            structure={},
            cells=[],
            metadata={},
        )
        result = self.engine.process(doc)
        assert "errors" in result
        assert len(result["errors"]) == 1
        assert "crasher" in result["errors"][0]

    def test_multiple_exceptions_all_recorded(self):
        """여러 플러그인에서 예외 발생 시 모두 기록"""
        bad1 = ExceptionPlugin("crasher1")
        bad2 = ExceptionPlugin("crasher2")
        good = FakePlugin("survivor")
        self.engine.register(bad1)
        self.engine.register(bad2)
        self.engine.register(good)

        doc = ParsedDocument(
            file_path="/tmp/test.xlsx",
            file_type="excel",
            raw_text="",
            structure={},
            cells=[],
            metadata={},
        )
        result = self.engine.process(doc)
        assert len(result["errors"]) == 2
        assert result.get("survivor") == "done"

    def test_process_returns_dict(self):
        """process()는 항상 dict를 반환해야 함"""
        doc = ParsedDocument(
            file_path="/tmp/test.xlsx",
            file_type="excel",
            raw_text="",
            structure={},
            cells=[],
            metadata={},
        )
        result = self.engine.process(doc)
        assert isinstance(result, dict)
