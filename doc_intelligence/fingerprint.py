"""
fingerprint.py — TF-IDF 핑거프린트 + 템플릿 매칭
Fingerprinter: 문서에서 핑거프린트 생성, 템플릿 학습/매칭
"""
import hashlib
import logging

try:
    import numpy as np
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    _ML_AVAILABLE = True
except ImportError as _e:
    np = None
    TfidfVectorizer = None
    cosine_similarity = None
    _ML_AVAILABLE = False
    logging.getLogger(__name__).warning(
        "numpy/scikit-learn 미인식 — 템플릿 매칭 비활성화 (%s). "
        "같은 인터프리터에 'python -m pip install numpy scikit-learn' 후 재시작.",
        _e,
    )

from doc_intelligence.engine import Fingerprint


class Fingerprinter:
    name = "fingerprinter"
    enabled = True

    def __init__(self, storage=None):
        self.storage = storage
        self._vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4)) if _ML_AVAILABLE else None
        self._corpus = []        # 학습된 텍스트 목록
        self._template_ids = []  # 대응 template ID

    def initialize(self, engine) -> None:
        """Engine에서 호출. 기존 템플릿을 DB에서 로드하고 vectorizer를 학습."""
        self.storage = engine.storage
        # DB에서 기존 템플릿 로드
        for t in self.storage.get_all_templates():
            label_positions = t["metadata"].get("label_positions", {})
            labels = list(label_positions.keys())
            self._corpus.append(" ".join(labels))
            self._template_ids.append(t["id"])
        if self._corpus and self._vectorizer is not None:
            self._vectorizer.fit(self._corpus)

    # ──────────────────────────────────────────────
    # 내부 유틸
    # ──────────────────────────────────────────────

    @staticmethod
    def _extract_labels(doc) -> dict:
        """셀에서 문자열 값만 추출하여 {값: 주소} 매핑 반환"""
        label_positions = {}
        for cell in doc.cells:
            if isinstance(cell.value, str) and cell.value.strip():
                label_positions[cell.value.strip()] = cell.address
        return label_positions

    @staticmethod
    def _merge_hash(doc) -> str:
        """structure.get('merge_cells', [])의 MD5 해시(32자 hex) 반환"""
        merge_cells = doc.structure.get("merge_cells", [])
        raw = str(sorted(merge_cells)).encode("utf-8")
        return hashlib.md5(raw).hexdigest()

    @staticmethod
    def _doc_id(doc) -> str:
        """file_path 기반 MD5 doc_id 생성"""
        return hashlib.md5(doc.file_path.encode("utf-8")).hexdigest()

    def _labels_to_text(self, label_positions: dict) -> str:
        """label_positions의 키를 공백 구분 텍스트로 변환"""
        return " ".join(label_positions.keys())

    def _vectorize(self, text: str) -> list:
        """단일 텍스트를 TF-IDF 벡터로 변환. corpus가 없거나 ML 미설치면 빈 리스트."""
        if not self._corpus or self._vectorizer is None:
            return []
        vec = self._vectorizer.transform([text])
        return vec.toarray()[0].tolist()

    # ──────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────

    def generate(self, doc) -> dict:
        """
        문서에서 핑거프린트를 생성한다.

        반환:
            {
                "labels": list[str],          # 추출된 라벨 목록
                "label_positions": dict,       # {값: 주소}
                "merge_pattern": str,          # MD5 해시
                "vector": list[float],         # TF-IDF 벡터
                "fingerprint": Fingerprint,    # dataclass 인스턴스
            }
        """
        label_positions = self._extract_labels(doc)
        labels = list(label_positions.keys())
        merge_pattern = self._merge_hash(doc)
        doc_id = self._doc_id(doc)
        text = self._labels_to_text(label_positions)
        vector = self._vectorize(text)

        fp = Fingerprint(
            doc_id=doc_id,
            feature_vector=vector,
            label_positions=label_positions,
            merge_pattern=merge_pattern,
        )

        return {
            "labels": labels,
            "label_positions": label_positions,
            "merge_pattern": merge_pattern,
            "vector": vector,
            "fingerprint": fp,
        }

    def learn(self, doc, template_name: str) -> int:
        """
        문서를 새 템플릿으로 학습하고 template_id를 반환한다.

        1. generate()로 핑거프린트 생성
        2. storage.save_template()으로 저장
        3. corpus에 추가, vectorizer 재학습
        """
        gen = self.generate(doc)
        label_positions = gen["label_positions"]
        merge_pattern = gen["merge_pattern"]

        metadata = {
            "label_positions": label_positions,
            "merge_pattern": merge_pattern,
        }
        fields = list(label_positions.keys())

        template_id = self.storage.save_template(
            name=template_name,
            fields=fields,
            metadata=metadata,
        )

        # corpus 및 vectorizer 갱신
        text = self._labels_to_text(label_positions)
        self._corpus.append(text)
        self._template_ids.append(template_id)
        if self._vectorizer is not None:
            self._vectorizer.fit(self._corpus)

        return template_id

    def match(self, doc) -> dict:
        """
        문서와 가장 유사한 템플릿을 찾아 반환한다.

        반환:
            {
                "template": int | None,   # template_id (없으면 None)
                "score": float,           # cosine similarity (0.0 ~ 1.0)
                "auto": bool,             # score >= 0.85 이면 True
            }

        임계값:
            >= 0.85 : auto=True,  storage.increment_match_count 호출
            0.60~0.84 : auto=False (후보 제시)
            < 0.60  : template=None
        """
        if not self._corpus or self._vectorizer is None:
            return {"template": None, "score": 0.0, "auto": False}

        label_positions = self._extract_labels(doc)
        text = self._labels_to_text(label_positions)

        # query + corpus를 합쳐서 TF-IDF 매트릭스 생성
        all_texts = self._corpus + [text]
        matrix = self._vectorizer.transform(all_texts)
        query_vec = matrix[-1]
        corpus_matrix = matrix[:-1]

        similarities = cosine_similarity(query_vec, corpus_matrix)[0]
        best_idx = int(np.argmax(similarities))
        best_score = float(similarities[best_idx])
        best_template_id = self._template_ids[best_idx]

        if best_score >= 0.85:
            self.storage.increment_match_count(best_template_id)
            return {"template": best_template_id, "score": best_score, "auto": True}
        elif best_score >= 0.60:
            return {"template": best_template_id, "score": best_score, "auto": False}
        else:
            return {"template": None, "score": best_score, "auto": False}

    def process(self, doc, context: dict) -> dict:
        """
        Engine 파이프라인에서 호출.
        context에 fingerprint, template_match 키를 추가한다.
        """
        context["fingerprint"] = self.generate(doc)
        context["template_match"] = self.match(doc)
        return context
