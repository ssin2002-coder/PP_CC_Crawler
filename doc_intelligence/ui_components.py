"""
ui_components.py — 6개 tkinter 위젯
LearningModeDialog, RuleManagerWidget, ValidationResultWidget,
OverlayWindow, DocumentListWidget, EntityListWidget
"""
import tkinter as tk
from tkinter import ttk


# ──────────────────────────────────────────────
# 1. LearningModeDialog
# ──────────────────────────────────────────────

class LearningModeDialog:
    """학습 모드 다이얼로그 — 엔티티별 필드 타입을 사용자가 지정한다."""

    FIELD_TYPES = [
        "날짜", "착공일", "준공일", "검수일", "금액", "예상비용", "부가세",
        "업체명", "부서", "설비코드", "부품코드", "이름", "검수자", "승인자",
        "문서번호", "문서 ID", "무시",
    ]

    def __init__(self, parent, entities: list, doc_name: str = ""):
        """
        parent  : tk.Tk 또는 tk.Toplevel (None 허용 — 테스트용)
        entities: [{"value": str, "type": str, ...}, ...]
        doc_name: 다이얼로그 제목에 표시할 문서명
        """
        self.entities = entities
        self.doc_name = doc_name
        self._corrections: dict = {}   # index -> 선택된 FIELD_TYPE
        self._combo_vars: list = []    # tk.StringVar 목록

        if parent is None:
            # 헤드리스/테스트 환경 — tkinter 없이 동작
            return

        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"학습 모드 — {doc_name}" if doc_name else "학습 모드")
        self.dialog.grab_set()

        self._build_ui()

    # ──────────────────────────────────────────────
    # 내부 UI 구성
    # ──────────────────────────────────────────────

    def _build_ui(self):
        frame = tk.Frame(self.dialog, padx=10, pady=10)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="엔티티", font=("맑은 고딕", 10, "bold")).grid(
            row=0, column=0, sticky="w", padx=4
        )
        tk.Label(frame, text="필드 타입", font=("맑은 고딕", 10, "bold")).grid(
            row=0, column=1, sticky="w", padx=4
        )

        for idx, entity in enumerate(self.entities):
            label_text = entity.get("value", "") if isinstance(entity, dict) else str(entity)
            tk.Label(frame, text=label_text).grid(
                row=idx + 1, column=0, sticky="w", padx=4, pady=2
            )

            default_type = entity.get("type", self.FIELD_TYPES[0]) if isinstance(entity, dict) else self.FIELD_TYPES[0]
            var = tk.StringVar(value=default_type)
            self._combo_vars.append(var)

            combo = ttk.Combobox(
                frame,
                textvariable=var,
                values=self.FIELD_TYPES,
                state="readonly",
                width=14,
            )
            combo.grid(row=idx + 1, column=1, padx=4, pady=2)

        btn_frame = tk.Frame(self.dialog)
        btn_frame.pack(pady=8)

        tk.Button(btn_frame, text="확인", command=self._on_ok, width=10).pack(
            side="left", padx=4
        )
        tk.Button(btn_frame, text="취소", command=self.dialog.destroy, width=10).pack(
            side="left", padx=4
        )

    def _on_ok(self):
        for idx, var in enumerate(self._combo_vars):
            self._corrections[idx] = var.get()
        self.dialog.destroy()

    # ──────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────

    def get_corrected_mappings(self) -> dict:
        """
        corrections 반영한 필드 매핑 반환.
        반환 형식: {index: field_type_str, ...}
        다이얼로그가 열리지 않은 경우(헤드리스)에는 각 엔티티의 기존 type을 그대로 사용.
        """
        if not self._corrections and self.entities:
            return {
                idx: (
                    entity.get("type", self.FIELD_TYPES[0])
                    if isinstance(entity, dict)
                    else self.FIELD_TYPES[0]
                )
                for idx, entity in enumerate(self.entities)
            }
        return dict(self._corrections)


# ──────────────────────────────────────────────
# 2. RuleManagerWidget
# ──────────────────────────────────────────────

class RuleManagerWidget:
    """규칙 관리 위젯 — 프리셋과 룰 목록을 표시하고 선택/편집한다."""

    def __init__(self, parent, presets: list, rules: list):
        """
        parent  : tk.Frame 또는 tk.Tk
        presets : [{"id": int, "name": str, ...}, ...]
        rules   : [{"id": int, "name": str, "rule_type": str, ...}, ...]
        """
        self.presets = presets
        self.rules = rules
        self._selected_preset = None
        self._selected_rule = None

        if parent is None:
            return

        self.frame = tk.Frame(parent, padx=8, pady=8)
        self.frame.pack(fill="both", expand=True)

        self._build_ui()

    # ──────────────────────────────────────────────
    # 내부 UI 구성
    # ──────────────────────────────────────────────

    def _build_ui(self):
        tk.Label(self.frame, text="프리셋", font=("맑은 고딕", 10, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        tk.Label(self.frame, text="룰 목록", font=("맑은 고딕", 10, "bold")).grid(
            row=0, column=1, sticky="w", padx=(16, 0)
        )

        # 프리셋 리스트박스
        self._preset_lb = tk.Listbox(self.frame, width=24, height=12, selectmode="single")
        self._preset_lb.grid(row=1, column=0, sticky="ns")
        for preset in self.presets:
            name = preset.get("name", "") if isinstance(preset, dict) else str(preset)
            self._preset_lb.insert("end", name)
        self._preset_lb.bind("<<ListboxSelect>>", self._on_preset_select)

        # 룰 리스트박스
        self._rule_lb = tk.Listbox(self.frame, width=28, height=12, selectmode="single")
        self._rule_lb.grid(row=1, column=1, sticky="ns", padx=(16, 0))
        for rule in self.rules:
            name = rule.get("name", "") if isinstance(rule, dict) else str(rule)
            self._rule_lb.insert("end", name)
        self._rule_lb.bind("<<ListboxSelect>>", self._on_rule_select)

    def _on_preset_select(self, event):
        selection = self._preset_lb.curselection()
        if selection:
            self._selected_preset = self.presets[selection[0]]

    def _on_rule_select(self, event):
        selection = self._rule_lb.curselection()
        if selection:
            self._selected_rule = self.rules[selection[0]]

    # ──────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────

    def get_selected_preset(self):
        return self._selected_preset

    def get_selected_rule(self):
        return self._selected_rule


# ──────────────────────────────────────────────
# 3. ValidationResultWidget
# ──────────────────────────────────────────────

class ValidationResultWidget:
    """검증 결과 위젯 — 룰별 통과/실패/경고를 테이블로 표시한다."""

    STATUS_COLORS = {
        "통과": "#28a745",
        "실패": "#dc3545",
        "경고": "#ffc107",
    }

    def __init__(self, parent, results: list):
        """
        parent  : tk.Frame 또는 tk.Tk
        results : [{"rule": str, "status": str, "detail": str}, ...]
        """
        self.results = results

        if parent is None:
            return

        self.frame = tk.Frame(parent, padx=8, pady=8)
        self.frame.pack(fill="both", expand=True)

        self._build_ui()

    # ──────────────────────────────────────────────
    # 내부 UI 구성
    # ──────────────────────────────────────────────

    def _build_ui(self):
        columns = ("rule", "status", "detail")
        self._tree = ttk.Treeview(
            self.frame,
            columns=columns,
            show="headings",
            height=12,
        )
        self._tree.heading("rule", text="룰 이름")
        self._tree.heading("status", text="상태")
        self._tree.heading("detail", text="상세")
        self._tree.column("rule", width=120)
        self._tree.column("status", width=70)
        self._tree.column("detail", width=300)

        for result in self.results:
            rule = result.get("rule", "")
            status = result.get("status", "")
            detail = result.get("detail", "")
            tag = status
            self._tree.insert("", "end", values=(rule, status, detail), tags=(tag,))

        for status, color in self.STATUS_COLORS.items():
            self._tree.tag_configure(status, foreground=color)

        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=scrollbar.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.frame.rowconfigure(0, weight=1)
        self.frame.columnconfigure(0, weight=1)


# ──────────────────────────────────────────────
# 4. OverlayWindow
# ──────────────────────────────────────────────

class OverlayWindow:
    """전체 화면 오버레이 — 드래그로 영역을 선택한다."""

    def __init__(self, on_region_selected=None):
        """
        on_region_selected: 영역 선택 완료 시 호출되는 콜백.
                            (x1, y1, x2, y2) 튜플을 인자로 받는다.
        """
        self.on_region_selected = on_region_selected
        self._start_x = 0
        self._start_y = 0
        self._rect_id = None

        try:
            self._win = tk.Toplevel()
            self._win.attributes("-fullscreen", True)
            self._win.attributes("-alpha", 0.3)
            self._win.configure(bg="black")
            self._win.attributes("-topmost", True)

            self._canvas = tk.Canvas(
                self._win,
                cursor="crosshair",
                bg="black",
                highlightthickness=0,
            )
            self._canvas.pack(fill="both", expand=True)

            self._canvas.bind("<ButtonPress-1>", self._on_press)
            self._canvas.bind("<B1-Motion>", self._on_drag)
            self._canvas.bind("<ButtonRelease-1>", self._on_release)
            self._win.bind("<Escape>", lambda e: self._win.destroy())

        except Exception:
            # 헤드리스 환경에서는 윈도우 생성 생략
            self._win = None
            self._canvas = None

    # ──────────────────────────────────────────────
    # 이벤트 핸들러
    # ──────────────────────────────────────────────

    def _on_press(self, event):
        self._start_x = event.x
        self._start_y = event.y
        if self._rect_id:
            self._canvas.delete(self._rect_id)
        self._rect_id = self._canvas.create_rectangle(
            self._start_x, self._start_y,
            self._start_x, self._start_y,
            outline="#58a6ff", width=2,
        )

    def _on_drag(self, event):
        if self._rect_id:
            self._canvas.coords(
                self._rect_id,
                self._start_x, self._start_y,
                event.x, event.y,
            )

    def _on_release(self, event):
        x1 = min(self._start_x, event.x)
        y1 = min(self._start_y, event.y)
        x2 = max(self._start_x, event.x)
        y2 = max(self._start_y, event.y)
        if self._win:
            self._win.destroy()
        if self.on_region_selected:
            self.on_region_selected(x1, y1, x2, y2)

    # ──────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────

    def close(self):
        if self._win:
            self._win.destroy()


# ──────────────────────────────────────────────
# 5. DocumentListWidget
# ──────────────────────────────────────────────

class DocumentListWidget:
    """열린 문서 목록 위젯 — 문서를 리스트로 표시한다."""

    def __init__(self, parent):
        self.documents: list = []

        if parent is None:
            return

        self.frame = tk.Frame(parent, padx=8, pady=8)
        self.frame.pack(fill="both", expand=True)

        self._build_ui()

    # ──────────────────────────────────────────────
    # 내부 UI 구성
    # ──────────────────────────────────────────────

    def _build_ui(self):
        tk.Label(self.frame, text="열린 문서", font=("맑은 고딕", 10, "bold")).pack(
            anchor="w"
        )
        self._listbox = tk.Listbox(self.frame, width=48, height=10, selectmode="single")
        self._listbox.pack(fill="both", expand=True)

        btn_frame = tk.Frame(self.frame)
        btn_frame.pack(pady=4)
        tk.Button(btn_frame, text="새로고침", command=self._refresh, width=10).pack(
            side="left", padx=4
        )

    def _refresh(self):
        self._listbox.delete(0, "end")
        for doc in self.documents:
            name = doc.get("name", "") if isinstance(doc, dict) else str(doc)
            self._listbox.insert("end", name)

    # ──────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────

    def add_document(self, doc_info: dict):
        self.documents.append(doc_info)
        if hasattr(self, "_listbox"):
            name = doc_info.get("name", "") if isinstance(doc_info, dict) else str(doc_info)
            self._listbox.insert("end", name)

    def clear(self):
        self.documents.clear()
        if hasattr(self, "_listbox"):
            self._listbox.delete(0, "end")


# ──────────────────────────────────────────────
# 6. EntityListWidget
# ──────────────────────────────────────────────

class EntityListWidget:
    """추출된 엔티티 목록 위젯 — 타입, 값, 위치, 신뢰도를 테이블로 표시한다."""

    def __init__(self, parent):
        self._entities: list = []

        if parent is None:
            return

        self.frame = tk.Frame(parent, padx=8, pady=8)
        self.frame.pack(fill="both", expand=True)

        self._build_ui()

    # ──────────────────────────────────────────────
    # 내부 UI 구성
    # ──────────────────────────────────────────────

    def _build_ui(self):
        tk.Label(self.frame, text="추출 엔티티", font=("맑은 고딕", 10, "bold")).pack(
            anchor="w"
        )
        columns = ("type", "value", "location", "confidence")
        self._tree = ttk.Treeview(
            self.frame,
            columns=columns,
            show="headings",
            height=10,
        )
        self._tree.heading("type", text="타입")
        self._tree.heading("value", text="값")
        self._tree.heading("location", text="위치")
        self._tree.heading("confidence", text="신뢰도")
        self._tree.column("type", width=80)
        self._tree.column("value", width=120)
        self._tree.column("location", width=120)
        self._tree.column("confidence", width=60)
        self._tree.pack(fill="both", expand=True)

    # ──────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────

    def set_entities(self, entities: list):
        """
        entities: Entity dataclass 목록 또는 dict 목록.
        기존 목록을 교체하고 트리뷰를 갱신한다.
        """
        self._entities = entities
        if not hasattr(self, "_tree"):
            return
        for item in self._tree.get_children():
            self._tree.delete(item)
        for entity in entities:
            if hasattr(entity, "type"):
                # Entity dataclass
                self._tree.insert(
                    "", "end",
                    values=(entity.type, entity.value, entity.location, f"{entity.confidence:.2f}"),
                )
            elif isinstance(entity, dict):
                self._tree.insert(
                    "", "end",
                    values=(
                        entity.get("type", ""),
                        entity.get("value", ""),
                        entity.get("location", ""),
                        entity.get("confidence", ""),
                    ),
                )

    def get_entities(self) -> list:
        return list(self._entities)
