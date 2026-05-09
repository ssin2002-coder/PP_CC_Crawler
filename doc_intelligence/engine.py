"""
engine.py — 코어 파이프라인 + 데이터 모델
데이터클래스: CellData, ParsedDocument, Entity, Fingerprint
Engine: 플러그인 등록/비활성화/파이프라인 실행
"""
import logging
from dataclasses import dataclass, field
from typing import List

from doc_intelligence.storage import Storage

logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────
# 데이터 모델
# ──────────────────────────────────────────────

@dataclass
class CellData:
    address: str        # "Sheet1!R2C3", "para:1", "slide1:shape1"
    value: object       # str, int, float 등
    data_type: str      # "text" | "number" | "date" | "formula"
    neighbors: dict     # 인접 셀 정보


@dataclass
class ParsedDocument:
    file_path: str
    file_type: str      # "excel" | "word" | "ppt" | "pdf" | "image"
    raw_text: str
    structure: dict
    cells: list         # list[CellData]
    metadata: dict


@dataclass
class Entity:
    type: str           # "금액" | "날짜" | "업체명" | "설비코드" 등
    value: str
    location: str
    confidence: float   # 0.0 ~ 1.0


@dataclass
class Fingerprint:
    doc_id: str
    feature_vector: list
    label_positions: dict
    merge_pattern: str


# ──────────────────────────────────────────────
# Engine
# ──────────────────────────────────────────────

class Engine:
    """플러그인 기반 문서 처리 파이프라인 엔진"""

    def __init__(self, db_path: str = "doc_intelligence.db"):
        self.storage = Storage(db_path=db_path)
        self.plugins: dict = {}      # name -> plugin 인스턴스
        self._order: list = []       # 등록 순서 보존
        self._disabled: set = set()  # 비활성화된 플러그인 이름

    def register(self, plugin) -> None:
        """플러그인을 등록하고 initialize()를 호출한다."""
        plugin.initialize(self)
        self.plugins[plugin.name] = plugin
        self._order.append(plugin.name)

    def disable(self, name: str) -> None:
        """플러그인을 비활성화한다. 존재하지 않는 이름은 무시한다."""
        self._disabled.add(name)

    def enable(self, name: str) -> None:
        """비활성화된 플러그인을 활성화한다. 존재하지 않는 이름은 무시한다."""
        self._disabled.discard(name)

    def list_plugins(self) -> List[str]:
        """등록된 플러그인 이름을 등록 순서대로 반환한다."""
        return list(self._order)

    def process(self, doc: ParsedDocument) -> dict:
        """
        등록 순서대로 활성화된 플러그인을 실행한다.
        각 플러그인의 결과는 context dict에 누적된다.
        플러그인 예외는 로깅 후 context["errors"]에 기록하고 다음 플러그인으로 진행한다.
        """
        context: dict = {"errors": []}

        for name in self._order:
            if name in self._disabled:
                continue

            plugin = self.plugins[name]
            try:
                context = plugin.process(doc, context)
                # process()가 errors 키를 제거했을 경우 복원
                if "errors" not in context:
                    context["errors"] = []
            except Exception as exc:
                error_msg = f"{name}: {exc}"
                logger.exception("플러그인 실행 중 예외 발생 — %s", error_msg)
                context["errors"].append(error_msg)

        return context
