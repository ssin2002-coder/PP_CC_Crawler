"""
가격 이력 관리 모듈
SQLite를 사용하여 자재별 단가 이력을 저장하고 통계를 계산합니다.
"""

import logging
import math
import os
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from typing import Dict, Generator, List, Optional

from backend.config import HISTORY_DIR, PRICE_HISTORY_DB

logger = logging.getLogger(__name__)


class HistoryManager:
    """
    자재 단가 이력을 SQLite 데이터베이스로 관리하는 클래스.

    첫 사용 시 테이블을 자동 생성합니다.
    """

    def __init__(self, db_path: Optional[str] = None) -> None:
        """
        Args:
            db_path: SQLite 데이터베이스 파일 경로.
                     None이면 config의 기본 경로(autosave/history/price_history.db)를 사용.
        """
        if db_path is None:
            db_path = os.path.join(HISTORY_DIR, PRICE_HISTORY_DB)

        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self._db_path = db_path
        self._init_db()

    def _init_db(self) -> None:
        """데이터베이스 테이블을 초기화합니다 (없으면 생성)."""
        with self._get_connection() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS price_history (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    material    TEXT    NOT NULL,
                    unit_price  REAL    NOT NULL,
                    date        TEXT    NOT NULL,
                    workbook    TEXT,
                    created_at  TEXT    DEFAULT (datetime('now', 'localtime'))
                )
            """)
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_material ON price_history (material)
            """)
            conn.commit()
        logger.debug(f"가격 이력 DB 초기화 완료: {self._db_path}")

    @contextmanager
    def _get_connection(self) -> Generator[sqlite3.Connection, None, None]:
        """SQLite 연결 컨텍스트 매니저입니다."""
        conn = sqlite3.connect(self._db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()

    def add_prices(self, records: List[Dict]) -> int:
        """
        가격 이력 레코드를 일괄 삽입합니다.

        Args:
            records: 레코드 목록. 각 레코드:
                {
                    'material': str,
                    'unit_price': float,
                    'date': str (ISO 형식 권장),
                    'workbook': str (선택)
                }

        Returns:
            int: 삽입된 레코드 수
        """
        if not records:
            return 0

        valid_records = []
        today = datetime.now().strftime('%Y-%m-%d')

        for rec in records:
            material = str(rec.get('material', '')).strip()
            unit_price = rec.get('unit_price')

            if not material or unit_price is None:
                logger.debug(f"유효하지 않은 레코드 스킵: {rec}")
                continue

            try:
                unit_price = float(unit_price)
            except (TypeError, ValueError):
                logger.debug(f"단가 변환 실패: {unit_price}")
                continue

            valid_records.append((
                material,
                unit_price,
                rec.get('date', today),
                rec.get('workbook', ''),
            ))

        if not valid_records:
            return 0

        with self._get_connection() as conn:
            conn.executemany(
                "INSERT INTO price_history (material, unit_price, date, workbook) VALUES (?, ?, ?, ?)",
                valid_records,
            )
            conn.commit()

        logger.info(f"가격 이력 {len(valid_records)}건 추가 완료")
        return len(valid_records)

    def get_stats(self, material: str) -> Optional[Dict]:
        """
        특정 자재의 단가 통계를 계산하여 반환합니다.

        Args:
            material: 자재 이름 (부분 일치 검색)

        Returns:
            dict | None: {mean, std, q1, q3, count, min, max}
                         데이터가 없으면 None
        """
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT unit_price FROM price_history WHERE material LIKE ? ORDER BY unit_price",
                (f'%{material}%',),
            ).fetchall()

        if not rows:
            return None

        values = [row['unit_price'] for row in rows]
        return self._compute_stats(values)

    def get_all_stats(self) -> Dict[str, Dict]:
        """
        모든 자재의 통계를 반환합니다.

        Returns:
            dict: {자재명: 통계 dict} 매핑
        """
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT material, unit_price FROM price_history ORDER BY material, unit_price"
            ).fetchall()

        # 자재별 그룹화
        groups: Dict[str, List[float]] = {}
        for row in rows:
            mat = row['material']
            if mat not in groups:
                groups[mat] = []
            groups[mat].append(row['unit_price'])

        result = {}
        for material, values in groups.items():
            result[material] = self._compute_stats(values)

        logger.debug(f"전체 통계 조회 완료: {len(result)}개 자재")
        return result

    def _compute_stats(self, values: List[float]) -> Dict:
        """
        숫자 목록의 통계를 계산합니다.

        Args:
            values: 정렬된 숫자 목록

        Returns:
            dict: {mean, std, q1, q3, count, min, max}
        """
        n = len(values)
        if n == 0:
            return {}

        total = sum(values)
        mean = total / n

        variance = sum((v - mean) ** 2 for v in values) / n
        std = math.sqrt(variance)

        sorted_vals = sorted(values)

        # Q1, Q3 계산
        q1_idx = n // 4
        q3_idx = (3 * n) // 4
        q1 = sorted_vals[q1_idx]
        q3 = sorted_vals[q3_idx]

        return {
            'mean': round(mean, 2),
            'std': round(std, 2),
            'q1': q1,
            'q3': q3,
            'count': n,
            'min': sorted_vals[0],
            'max': sorted_vals[-1],
        }

    def get_materials(self) -> List[str]:
        """
        저장된 모든 자재 이름 목록을 반환합니다.

        Returns:
            List[str]: 자재 이름 목록 (가나다 순)
        """
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT DISTINCT material FROM price_history ORDER BY material"
            ).fetchall()
        return [row['material'] for row in rows]


# 전역 싱글턴 인스턴스
_history_manager: Optional[HistoryManager] = None


def get_history_manager() -> HistoryManager:
    """HistoryManager 싱글턴 인스턴스를 반환합니다."""
    global _history_manager
    if _history_manager is None:
        _history_manager = HistoryManager()
    return _history_manager
