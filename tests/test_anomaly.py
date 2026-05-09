"""
test_anomaly.py — AnomalyDetector 단위 테스트
테스트 1: detect_outlier — 이상치 탐지 (마지막 값이 True)
테스트 2: no_outlier — 균등 값에서 이상치 없음
테스트 3: insufficient_data — 데이터 < 5이면 모두 False
"""
import pytest
from doc_intelligence.anomaly import AnomalyDetector


class TestAnomalyDetector:

    def test_detect_outlier(self):
        """[100, 110, 105, 95, 500] — 마지막 값(500)이 이상치로 탐지되어야 함"""
        detector = AnomalyDetector(contamination=0.1)
        values = [100, 110, 105, 95, 500]
        result = detector.detect(values)

        assert len(result) == 5
        # 마지막 값(500)이 이상치
        assert result[-1] == True

    def test_no_outlier(self):
        """균등한 값 [100, 100, 100, 100, 100] — 이상치 없음"""
        detector = AnomalyDetector(contamination=0.05)
        values = [100, 100, 100, 100, 100]
        result = detector.detect(values)

        assert len(result) == 5
        # 모두 정상 (이상치 없음)
        assert not any(result)

    def test_insufficient_data(self):
        """데이터 길이 < 5이면 모두 False 반환"""
        detector = AnomalyDetector()
        values = [100, 200, 300]
        result = detector.detect(values)

        assert len(result) == 3
        assert not any(result)
