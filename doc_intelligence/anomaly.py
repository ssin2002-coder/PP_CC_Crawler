"""
anomaly.py — Isolation Forest 기반 이상 탐지
AnomalyDetector: 금액 엔티티에서 수치 추출 후 이상치 판별
"""
import re
import numpy as np
from sklearn.ensemble import IsolationForest


class AnomalyDetector:
    name = "anomaly_detector"
    enabled = True

    def __init__(self, contamination=0.1):
        self.contamination = contamination

    def initialize(self, engine):
        pass

    def detect(self, values) -> list:
        """
        values: 수치 리스트
        len < 5이면 [False] * len(values) 반환
        IsolationForest(contamination, random_state=42).fit_predict 사용
        -1이면 True(이상치), 1이면 False(정상)
        """
        if len(values) < 5:
            return [False] * len(values)

        X = np.array(values).reshape(-1, 1)
        model = IsolationForest(
            contamination=self.contamination,
            random_state=42,
        )
        preds = model.fit_predict(X)
        return [bool(pred == -1) for pred in preds]

    def process(self, doc, context):
        """
        context["entities"]에서 금액 엔티티를 추출하여
        수치 파싱 후 detect() 호출, 결과를 context["anomalies"]에 저장.
        anomalies: [{"entity": Entity, "is_anomaly": bool}, ...]
        """
        entities = context.get("entities", [])
        amount_entities = [e for e in entities if e.type == "금액"]

        values = []
        for entity in amount_entities:
            # "원" 제거 후 쉼표 제거하여 수치 파싱
            raw = re.sub(r"[^\d]", "", entity.value)
            if raw:
                values.append(float(raw))
            else:
                values.append(0.0)

        flags = self.detect(values)

        anomalies = []
        for entity, is_anomaly in zip(amount_entities, flags):
            anomalies.append({
                "entity": entity,
                "is_anomaly": is_anomaly,
            })

        context["anomalies"] = anomalies
        return context
