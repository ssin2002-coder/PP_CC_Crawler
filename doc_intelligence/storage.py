"""
storage.py — SQLite 기반 저장소
테이블: templates, rules, presets, documents, validation_results
"""
import json
import sqlite3
from typing import Any, Dict, List, Optional


class Storage:
    """SQLite 전체 CRUD 저장소"""

    def __init__(self, db_path: str = "doc_intelligence.db"):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self._init_tables()

    # ──────────────────────────────────────────────
    # 내부 유틸
    # ──────────────────────────────────────────────
    def _init_tables(self) -> None:
        """5개 테이블 초기화"""
        cursor = self.conn.cursor()
        cursor.executescript("""
            CREATE TABLE IF NOT EXISTS templates (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                name        TEXT    NOT NULL,
                fields      TEXT    NOT NULL DEFAULT '[]',
                metadata    TEXT    NOT NULL DEFAULT '{}',
                match_count INTEGER NOT NULL DEFAULT 0,
                created_at  TEXT    NOT NULL DEFAULT (datetime('now', 'localtime')),
                updated_at  TEXT    NOT NULL DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS rules (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                name        TEXT    NOT NULL,
                rule_type   TEXT    NOT NULL,
                conditions  TEXT    NOT NULL DEFAULT '{}',
                actions     TEXT    NOT NULL DEFAULT '{}',
                created_at  TEXT    NOT NULL DEFAULT (datetime('now', 'localtime')),
                updated_at  TEXT    NOT NULL DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS presets (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                name         TEXT    NOT NULL,
                template_ids TEXT    NOT NULL DEFAULT '[]',
                rule_ids     TEXT    NOT NULL DEFAULT '[]',
                settings     TEXT    NOT NULL DEFAULT '{}',
                created_at   TEXT    NOT NULL DEFAULT (datetime('now', 'localtime')),
                updated_at   TEXT    NOT NULL DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS documents (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                filename    TEXT    NOT NULL,
                filepath    TEXT    NOT NULL,
                template_id INTEGER,
                parsed_data TEXT    NOT NULL DEFAULT '{}',
                created_at  TEXT    NOT NULL DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS validation_results (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                preset_id    INTEGER,
                rule_id      INTEGER,
                document_ids TEXT    NOT NULL DEFAULT '[]',
                status       TEXT    NOT NULL,
                detail       TEXT    NOT NULL DEFAULT '{}',
                created_at   TEXT    NOT NULL DEFAULT (datetime('now', 'localtime')),
                FOREIGN KEY (preset_id) REFERENCES presets(id)
            );
        """)
        self.conn.commit()

    @staticmethod
    def _dumps(obj: Any) -> str:
        return json.dumps(obj, ensure_ascii=False)

    @staticmethod
    def _loads(s: str) -> Any:
        return json.loads(s)

    def _row_to_dict(self, row) -> Optional[Dict]:
        if row is None:
            return None
        return dict(row)

    # ──────────────────────────────────────────────
    # 기타
    # ──────────────────────────────────────────────
    def list_tables(self) -> List[str]:
        """현재 DB에 존재하는 테이블 이름 목록 반환"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
        return [row["name"] for row in cursor.fetchall()]

    def close(self) -> None:
        """DB 연결 종료"""
        if self.conn:
            self.conn.close()

    # ──────────────────────────────────────────────
    # Templates
    # ──────────────────────────────────────────────
    def save_template(
        self,
        name: str,
        fields: List[str],
        metadata: Dict
    ) -> int:
        """템플릿 저장 후 ID 반환"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT INTO templates (name, fields, metadata)
            VALUES (?, ?, ?)
            """,
            (name, self._dumps(fields), self._dumps(metadata))
        )
        self.conn.commit()
        return cursor.lastrowid

    def get_template(self, template_id: int) -> Optional[Dict]:
        """ID로 템플릿 조회. 없으면 None 반환"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM templates WHERE id = ?", (template_id,))
        row = cursor.fetchone()
        if row is None:
            return None
        d = dict(row)
        d["fields"] = self._loads(d["fields"])
        d["metadata"] = self._loads(d["metadata"])
        return d

    def get_all_templates(self) -> List[Dict]:
        """전체 템플릿 목록 반환"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM templates ORDER BY id")
        rows = cursor.fetchall()
        result = []
        for row in rows:
            d = dict(row)
            d["fields"] = self._loads(d["fields"])
            d["metadata"] = self._loads(d["metadata"])
            result.append(d)
        return result

    def update_template(self, template_id: int, **kwargs) -> None:
        """템플릿 수정. kwargs로 유연하게 처리"""
        if not kwargs:
            return
        json_fields = {"fields", "metadata"}
        set_parts = []
        values = []
        for key, value in kwargs.items():
            set_parts.append(f"{key} = ?")
            values.append(self._dumps(value) if key in json_fields else value)
        set_parts.append("updated_at = datetime('now', 'localtime')")
        values.append(template_id)
        sql = f"UPDATE templates SET {', '.join(set_parts)} WHERE id = ?"
        self.conn.execute(sql, values)
        self.conn.commit()

    def delete_template(self, template_id: int) -> None:
        """템플릿 삭제"""
        self.conn.execute("DELETE FROM templates WHERE id = ?", (template_id,))
        self.conn.commit()

    def increment_match_count(self, template_id: int) -> None:
        """match_count 1 증가"""
        self.conn.execute(
            "UPDATE templates SET match_count = match_count + 1 WHERE id = ?",
            (template_id,)
        )
        self.conn.commit()

    # ──────────────────────────────────────────────
    # Rules
    # ──────────────────────────────────────────────
    def save_rule(
        self,
        name: str,
        rule_type: str,
        conditions: Dict,
        actions: Dict
    ) -> int:
        """규칙 저장 후 ID 반환"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT INTO rules (name, rule_type, conditions, actions)
            VALUES (?, ?, ?, ?)
            """,
            (name, rule_type, self._dumps(conditions), self._dumps(actions))
        )
        self.conn.commit()
        return cursor.lastrowid

    def get_rule(self, rule_id: int) -> Optional[Dict]:
        """ID로 규칙 조회. 없으면 None 반환"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM rules WHERE id = ?", (rule_id,))
        row = cursor.fetchone()
        if row is None:
            return None
        d = dict(row)
        d["conditions"] = self._loads(d["conditions"])
        d["actions"] = self._loads(d["actions"])
        return d

    def get_all_rules(self) -> List[Dict]:
        """전체 규칙 목록 반환"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM rules ORDER BY id")
        rows = cursor.fetchall()
        result = []
        for row in rows:
            d = dict(row)
            d["conditions"] = self._loads(d["conditions"])
            d["actions"] = self._loads(d["actions"])
            result.append(d)
        return result

    def update_rule(self, rule_id: int, **kwargs) -> None:
        """규칙 수정. kwargs로 유연하게 처리"""
        if not kwargs:
            return
        json_fields = {"conditions", "actions"}
        set_parts = []
        values = []
        for key, value in kwargs.items():
            set_parts.append(f"{key} = ?")
            values.append(self._dumps(value) if key in json_fields else value)
        set_parts.append("updated_at = datetime('now', 'localtime')")
        values.append(rule_id)
        sql = f"UPDATE rules SET {', '.join(set_parts)} WHERE id = ?"
        self.conn.execute(sql, values)
        self.conn.commit()

    def delete_rule(self, rule_id: int) -> None:
        """규칙 삭제"""
        self.conn.execute("DELETE FROM rules WHERE id = ?", (rule_id,))
        self.conn.commit()

    # ──────────────────────────────────────────────
    # Presets
    # ──────────────────────────────────────────────
    def save_preset(
        self,
        name: str,
        template_ids: List[int],
        rule_ids: List[int],
        settings: Dict
    ) -> int:
        """프리셋 저장 후 ID 반환"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT INTO presets (name, template_ids, rule_ids, settings)
            VALUES (?, ?, ?, ?)
            """,
            (
                name,
                self._dumps(template_ids),
                self._dumps(rule_ids),
                self._dumps(settings)
            )
        )
        self.conn.commit()
        return cursor.lastrowid

    def get_preset(self, preset_id: int) -> Optional[Dict]:
        """ID로 프리셋 조회. 없으면 None 반환"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM presets WHERE id = ?", (preset_id,))
        row = cursor.fetchone()
        if row is None:
            return None
        d = dict(row)
        d["template_ids"] = self._loads(d["template_ids"])
        d["rule_ids"] = self._loads(d["rule_ids"])
        d["settings"] = self._loads(d["settings"])
        return d

    def get_all_presets(self) -> List[Dict]:
        """전체 프리셋 목록 반환"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM presets ORDER BY id")
        rows = cursor.fetchall()
        result = []
        for row in rows:
            d = dict(row)
            d["template_ids"] = self._loads(d["template_ids"])
            d["rule_ids"] = self._loads(d["rule_ids"])
            d["settings"] = self._loads(d["settings"])
            result.append(d)
        return result

    def update_preset(self, preset_id: int, **kwargs) -> None:
        """프리셋 수정. kwargs로 유연하게 처리"""
        if not kwargs:
            return
        json_fields = {"template_ids", "rule_ids", "settings"}
        set_parts = []
        values = []
        for key, value in kwargs.items():
            set_parts.append(f"{key} = ?")
            values.append(self._dumps(value) if key in json_fields else value)
        set_parts.append("updated_at = datetime('now', 'localtime')")
        values.append(preset_id)
        sql = f"UPDATE presets SET {', '.join(set_parts)} WHERE id = ?"
        self.conn.execute(sql, values)
        self.conn.commit()

    def delete_preset(self, preset_id: int) -> None:
        """프리셋 삭제"""
        self.conn.execute("DELETE FROM presets WHERE id = ?", (preset_id,))
        self.conn.commit()

    def find_presets_by_template_ids(self, open_template_ids: List[int]) -> List[Dict]:
        """
        열린 문서들의 template_id 조합이 프리셋의 template_ids의 부분집합인지 확인.

        매칭 조건:
          open_template_ids ⊆ preset.template_ids
          즉, 열린 문서 ID가 모두 프리셋 목록 안에 있으면 매칭.
          열린 문서 중 프리셋에 없는 항목이 하나라도 있으면 미매칭.
        """
        open_set = set(open_template_ids)
        matched = []
        for preset in self.get_all_presets():
            preset_set = set(preset["template_ids"])
            if open_set.issubset(preset_set):
                matched.append(preset)
        return matched

    # ──────────────────────────────────────────────
    # Documents
    # ──────────────────────────────────────────────
    def save_document(
        self,
        filename: str,
        filepath: str,
        template_id: Optional[int],
        parsed_data: Dict
    ) -> int:
        """문서 저장 후 ID 반환"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT INTO documents (filename, filepath, template_id, parsed_data)
            VALUES (?, ?, ?, ?)
            """,
            (filename, filepath, template_id, self._dumps(parsed_data))
        )
        self.conn.commit()
        return cursor.lastrowid

    # ──────────────────────────────────────────────
    # Validation Results
    # ──────────────────────────────────────────────
    def save_validation_result(
        self,
        preset_id: Optional[int],
        rule_id: Optional[int],
        document_ids: List[int],
        status: str,
        detail: Dict
    ) -> int:
        """
        검증 결과 저장 후 ID 반환.
        status 값: "통과", "실패", "경고"
        document_ids: 검증 대상 문서 ID 목록 (JSON 배열로 저장)
        """
        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT INTO validation_results (preset_id, rule_id, document_ids, status, detail)
            VALUES (?, ?, ?, ?, ?)
            """,
            (preset_id, rule_id, self._dumps(document_ids), status, self._dumps(detail))
        )
        self.conn.commit()
        return cursor.lastrowid

    def get_validation_results(self, preset_id: Optional[int] = None) -> List[Dict]:
        """
        검증 결과 목록 반환.
        preset_id 지정 시 해당 프리셋 결과만 필터링. None이면 전체 반환.
        document_ids는 json.loads로 파싱하여 반환.
        """
        cursor = self.conn.cursor()
        if preset_id is not None:
            cursor.execute(
                "SELECT * FROM validation_results WHERE preset_id = ? ORDER BY id",
                (preset_id,)
            )
        else:
            cursor.execute("SELECT * FROM validation_results ORDER BY id")
        rows = cursor.fetchall()
        result = []
        for row in rows:
            d = dict(row)
            d["document_ids"] = self._loads(d["document_ids"])
            d["detail"] = self._loads(d["detail"])
            result.append(d)
        return result
