"""
extractor.py — regex + 인접 셀 기반 엔티티 추출기
PATTERNS 매칭으로 Entity 생성, LABEL_HINTS 인접 셀로 confidence 보정,
template["field_mappings"] 기반 정확 추출 지원
"""
import re
from doc_intelligence.engine import Entity


class EntityExtractor:
    name = "extractor"
    enabled = True

    PATTERNS = {
        "금액": r'[\d,]+\s*원',
        "날짜": r'\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}',
        "설비코드": r'[A-Z]{2,3}-\d{3,5}',
        "사업자번호": r'\d{3}-\d{2}-\d{5}',
        "전화번호": r'\d{2,3}-\d{3,4}-\d{4}',
    }

    LABEL_HINTS = {
        "합계": "금액", "금액": "금액", "비용": "금액", "단가": "금액", "소계": "금액",
        "날짜": "날짜", "일자": "날짜", "착공": "날짜", "준공": "날짜", "검수": "날짜",
        "업체": "업체명", "시공": "업체명", "업체명": "업체명",
        "설비": "설비코드", "장비": "설비코드",
    }

    def initialize(self, engine):
        pass

    def extract(self, doc, template=None) -> list:
        """
        template 제공 시 field_mappings 기반 정확 추출,
        없으면 전체 셀 자동 추출 후 인접 셀 confidence 보정.
        """
        if template:
            return self._extract_by_template(doc, template)
        return self._extract_auto(doc)

    def _extract_auto(self, doc):
        """전체 셀에서 PATTERNS 매칭 -> Entity 생성 (confidence=0.7), 인접 셀 힌트로 보정."""
        cells = doc.cells
        entities = []

        for cell in cells:
            cell_str = str(cell.value) if cell.value is not None else ""
            for entity_type, pattern in self.PATTERNS.items():
                match = re.search(pattern, cell_str)
                if match:
                    entities.append(Entity(
                        type=entity_type,
                        value=match.group(),
                        location=cell.address,
                        confidence=0.7,
                    ))

        entities = self._boost_by_neighbors(cells, entities)
        return entities

    def _boost_by_neighbors(self, cells, entities):
        """
        LABEL_HINTS 키워드가 있는 셀의 바로 다음 셀이 엔티티면 confidence += 0.2.
        결과 confidence는 최대 1.0으로 제한.
        """
        # address -> entity 매핑 (빠른 조회용)
        location_map: dict = {}
        for entity in entities:
            location_map.setdefault(entity.location, []).append(entity)

        for i, cell in enumerate(cells):
            cell_str = str(cell.value) if cell.value is not None else ""
            # LABEL_HINTS 키워드 포함 여부 확인
            for hint_key in self.LABEL_HINTS:
                if hint_key in cell_str:
                    # 다음 셀이 존재하면
                    if i + 1 < len(cells):
                        next_cell = cells[i + 1]
                        if next_cell.address in location_map:
                            for entity in location_map[next_cell.address]:
                                entity.confidence = min(1.0, entity.confidence + 0.2)
                    break

        return entities

    def _extract_by_template(self, doc, template):
        """
        template["field_mappings"]의 주소 기반 정확 추출 (confidence=0.95).
        field_mappings: { field_name: { "address": "Sheet1!R2C3", "type": "금액" } }
        """
        field_mappings = template.get("field_mappings", {})
        if not field_mappings:
            return []

        # address -> cell 매핑
        cell_map = {cell.address: cell for cell in doc.cells}

        entities = []
        for field_name, mapping in field_mappings.items():
            address = mapping.get("address", "")
            entity_type = mapping.get("type", "unknown")
            cell = cell_map.get(address)
            if cell is not None:
                cell_str = str(cell.value) if cell.value is not None else ""
                entities.append(Entity(
                    type=entity_type,
                    value=cell_str,
                    location=address,
                    confidence=0.95,
                ))

        return entities

    def process(self, doc, context):
        """Engine 파이프라인 진입점. template_match 결과를 활용해 추출 후 context에 저장."""
        template = context.get("template_match", {}).get("template")
        entities = self.extract(doc, template)
        context["entities"] = entities
        return context
