"""
graph.py — NetworkX 기반 문서 관계 그래프
DocGraph: 문서를 노드, 검증 규칙 결과를 엣지로 시각화
"""
import networkx as nx


# 검증 상태별 엣지 색상 매핑
STATUS_COLORS = {
    "pass": "green",
    "fail": "red",
    "warning": "orange",
    "unknown": "gray",
}


class DocGraph:
    name = "doc_graph"
    enabled = True

    def __init__(self):
        self.G = nx.Graph()

    def initialize(self, engine):
        pass

    def add_document(self, doc_name, entities):
        """문서를 노드로 추가한다."""
        self.G.add_node(doc_name, entities=entities)

    def add_relationship(self, doc1, doc2, rule_name, status):
        """
        두 문서 간 관계를 엣지로 추가한다.
        status에 따라 color 속성을 매핑한다.
        """
        color = STATUS_COLORS.get(status, STATUS_COLORS["unknown"])
        self.G.add_edge(doc1, doc2, rule=rule_name, status=status, color=color)

    def node_count(self):
        """등록된 노드 수를 반환한다."""
        return self.G.number_of_nodes()

    def get_edges(self):
        """모든 엣지를 (node1, node2, data_dict) 형태로 반환한다."""
        return list(self.G.edges(data=True))

    def to_html(self):
        """
        pyvis.network.Network로 HTML을 생성한다.
        pyvis를 import할 수 없으면 간단한 HTML을 반환한다.
        """
        try:
            from pyvis.network import Network

            net = Network(height="600px", width="100%", bgcolor="#222222", font_color="white")
            net.from_nx(self.G)

            # 엣지 색상 적용
            for edge in net.edges:
                src = edge.get("from", "")
                dst = edge.get("to", "")
                edge_data = self.G.get_edge_data(src, dst) or {}
                edge["color"] = edge_data.get("color", "gray")

            html = net.generate_html()
            return html

        except ImportError:
            # pyvis 없을 경우 간단한 fallback HTML
            nodes_html = "".join(
                f"<li>{node}</li>" for node in self.G.nodes()
            )
            edges_html = "".join(
                f"<li>{u} — {v} [{data.get('rule', '')} / {data.get('status', '')}]</li>"
                for u, v, data in self.G.edges(data=True)
            )
            return (
                "<!DOCTYPE html><html><body>"
                f"<h2>Document Graph</h2>"
                f"<h3>Nodes</h3><ul>{nodes_html}</ul>"
                f"<h3>Edges</h3><ul>{edges_html}</ul>"
                "</body></html>"
            )

    def process(self, doc, context):
        """
        Engine 파이프라인 진입점.
        doc을 노드로 추가하고 context["graph"]에 self를 저장한다.
        """
        entities = context.get("entities", [])
        self.add_document(doc.file_path, entities)
        context["graph"] = self
        return context
