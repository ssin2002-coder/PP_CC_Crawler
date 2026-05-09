"""
test_graph.py — DocGraph 단위 테스트
테스트 1: add_documents — 노드 추가 후 node_count 확인
테스트 2: add_relationship — 엣지 추가 및 속성(rule, status, color) 확인
테스트 3: to_html — HTML 문자열 반환 확인
"""
import pytest
from doc_intelligence.graph import DocGraph


class TestDocGraph:

    def test_add_documents(self):
        """노드 2개 추가 후 node_count == 2"""
        graph = DocGraph()
        graph.add_document("doc_a.xlsx", [])
        graph.add_document("doc_b.docx", ["entity1"])

        assert graph.node_count() == 2

    def test_add_relationship(self):
        """엣지 추가 후 rule, status, color 속성 확인"""
        graph = DocGraph()
        graph.add_document("doc_a.xlsx", [])
        graph.add_document("doc_b.docx", [])
        graph.add_relationship("doc_a.xlsx", "doc_b.docx", "amount_match", "pass")

        edges = graph.get_edges()
        assert len(edges) == 1

        _, _, data = edges[0]
        assert data["rule"] == "amount_match"
        assert data["status"] == "pass"
        assert data["color"] == "green"

    def test_to_html(self):
        """to_html()이 HTML 문자열을 반환해야 함"""
        graph = DocGraph()
        graph.add_document("doc_a.xlsx", [])
        graph.add_document("doc_b.docx", [])
        graph.add_relationship("doc_a.xlsx", "doc_b.docx", "date_match", "fail")

        html = graph.to_html()

        assert isinstance(html, str)
        assert len(html) > 0
        # HTML 태그 포함 여부 확인
        assert "<" in html and ">" in html
