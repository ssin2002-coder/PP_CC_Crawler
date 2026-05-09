"""
main.py — Doc Intelligence 메인 앱
DocIntelligenceApp: COM 폴링 + 파이프라인 실행
main(): tkinter GUI 5탭 다크 테마 또는 터미널 fallback 실행
"""
import logging
import queue
import threading
import time
from datetime import datetime

from doc_intelligence.engine import Engine
from doc_intelligence.com_worker import ComWorker
from doc_intelligence.parsers import ExcelParser, WordParser, PowerPointParser, PdfParser, ImageParser
from doc_intelligence.fingerprint import Fingerprinter
from doc_intelligence.extractor import EntityExtractor

logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────
# COM ProgID → 파서 키 매핑
# ──────────────────────────────────────────────

_APP_TO_PARSER = {
    "Excel.Application": "excel",
    "Word.Application": "word",
    "PowerPoint.Application": "ppt",
    "AcroExch.App": "pdf",
}

# 문서 타입별 아이콘 레이블
_APP_TO_ICON = {
    "excel": "XL",
    "word": "W",
    "ppt": "PPT",
    "pdf": "PDF",
    "image": "IMG",
}


class DocIntelligenceApp:
    """COM 폴링 기반 문서 처리 앱."""

    def __init__(self):
        self.engine = Engine(db_path="templates.db")
        self.com_worker = ComWorker()
        self.parsers = {
            "excel": ExcelParser(),
            "word": WordParser(),
            "ppt": PowerPointParser(),
            "pdf": PdfParser(),
            "image": ImageParser(),
        }
        self.engine.register(Fingerprinter())
        self.engine.register(EntityExtractor())
        self._polling = False
        self._poll_thread: threading.Thread | None = None
        # 이미 처리한 문서 경로 추적 (중복 처리 방지)
        self._seen_paths: set = set()
        # 열린 문서 목록
        self.open_docs: list = []
        # 추출된 엔티티 목록
        self.entities: list = []
        # 검증 결과 목록
        self.validation_results: list = []
        # 활동 로그
        self.activity_log: list = []
        # UI 이벤트 큐 (스레드 → 메인 스레드)
        self._ui_queue: queue.Queue = queue.Queue()
        # UI 콜백 (main()에서 설정)
        self.on_ui_update = None

    # ──────────────────────────────────────────────
    # 폴링
    # ──────────────────────────────────────────────

    def start_polling(self, interval: int = 3) -> None:
        """별도 스레드에서 com_worker.detect_open_documents 폴링.
        새 문서 감지 시 _process_document 호출.
        """
        self._polling = True
        self._poll_thread = threading.Thread(
            target=self._poll_loop,
            args=(interval,),
            daemon=True,
            name="DocIntelligencePollThread",
        )
        self._poll_thread.start()
        logger.info("COM 폴링 시작 (interval=%ds)", interval)

    def stop_polling(self) -> None:
        """폴링을 중지한다."""
        self._polling = False
        if self._poll_thread and self._poll_thread.is_alive():
            self._poll_thread.join(timeout=5)
        logger.info("COM 폴링 중지")

    def _poll_loop(self, interval: int) -> None:
        """폴링 루프 — _polling이 True인 동안 반복 실행한다."""
        while self._polling:
            try:
                docs = self.com_worker.detect_open_documents()
                for doc_info in docs:
                    path = doc_info.get("path", "")
                    if path and path not in self._seen_paths:
                        self._seen_paths.add(path)
                        try:
                            context = self._process_document(doc_info)
                            # UI 큐에 이벤트 전송
                            self._ui_queue.put(("doc_processed", doc_info, context))
                        except Exception as exc:
                            logger.exception("문서 처리 중 예외 — %s: %s", path, exc)
            except Exception as exc:
                logger.exception("COM 감지 중 예외: %s", exc)
            time.sleep(interval)

    # ──────────────────────────────────────────────
    # 문서 처리
    # ──────────────────────────────────────────────

    def _process_document(self, doc_info: dict) -> dict:
        """
        parser 선택 → COM 파싱 → engine.process → 결과 반환.

        doc_info: {"app": prog_id, "name": 문서명, "path": 전체경로}
        반환: engine.process context dict
        """
        app = doc_info.get("app", "")
        parser_key = _APP_TO_PARSER.get(app)

        if parser_key is None:
            logger.warning("지원하지 않는 앱 ProgID: %s", app)
            return {}

        parser = self.parsers.get(parser_key)
        if parser is None:
            logger.warning("파서를 찾을 수 없음: %s", parser_key)
            return {}

        com_app = self.com_worker.get_active_app(app)
        parsed_doc = parser.parse_from_com(com_app)
        context = self.engine.process(parsed_doc)
        logger.info("문서 처리 완료 — %s / 엔티티 %d개",
                    doc_info.get("name", ""), len(context.get("entities", [])))
        return context

    def add_activity_log(self, message: str) -> None:
        """활동 로그에 항목을 추가한다."""
        now = datetime.now().strftime("%H:%M")
        self.activity_log.append({"time": now, "message": message})
        if len(self.activity_log) > 100:
            self.activity_log = self.activity_log[-100:]


# ──────────────────────────────────────────────
# 다크 테마 색상 상수
# ──────────────────────────────────────────────

COLORS = {
    "bg_main":    "#0f1117",
    "bg_panel":   "#161b22",
    "bg_card":    "#21262d",
    "bg_dark":    "#0d1117",
    "border":     "#21262d",
    "border2":    "#30363d",
    "text_main":  "#e0e0e0",
    "text_sub":   "#8b949e",
    "text_head":  "#c9d1d9",
    "accent":     "#58a6ff",
    "green":      "#3fb950",
    "red":        "#f85149",
    "yellow":     "#d29922",
    "bg_green":   "#1a3d1a",
    "bg_red":     "#3d1a1a",
    "bg_yellow":  "#3d2e00",
    "bg_blue":    "#1a2d3d",
    "btn_bg":     "#21262d",
    "btn_hover":  "#30363d",
    "btn_primary":"#1f6feb",
    "btn_success":"#238636",
}

FONTS = {
    "title":   ("맑은 고딕", 13, "bold"),
    "normal":  ("맑은 고딕", 10),
    "small":   ("맑은 고딕", 9),
    "bold":    ("맑은 고딕", 10, "bold"),
    "mono":    ("Consolas", 9),
}


# ──────────────────────────────────────────────
# GUI 빌더 함수
# ──────────────────────────────────────────────

def _apply_dark_style(style) -> None:
    """ttk 위젯에 다크 테마 스타일을 적용한다."""
    style.theme_use("clam")
    style.configure(
        ".",
        background=COLORS["bg_main"],
        foreground=COLORS["text_main"],
        fieldbackground=COLORS["bg_panel"],
        troughcolor=COLORS["bg_card"],
        selectbackground=COLORS["btn_primary"],
        selectforeground="#ffffff",
        bordercolor=COLORS["border2"],
        darkcolor=COLORS["bg_dark"],
        lightcolor=COLORS["bg_panel"],
    )
    style.configure("TNotebook", background=COLORS["bg_panel"], borderwidth=0)
    style.configure(
        "TNotebook.Tab",
        background=COLORS["bg_card"],
        foreground=COLORS["text_sub"],
        padding=[14, 7],
        borderwidth=0,
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", COLORS["bg_main"]), ("active", COLORS["bg_panel"])],
        foreground=[("selected", COLORS["accent"]), ("active", COLORS["text_head"])],
    )
    style.configure(
        "Treeview",
        background=COLORS["bg_dark"],
        foreground=COLORS["text_main"],
        fieldbackground=COLORS["bg_dark"],
        rowheight=28,
        borderwidth=0,
    )
    style.configure(
        "Treeview.Heading",
        background=COLORS["bg_panel"],
        foreground=COLORS["text_head"],
        borderwidth=0,
        relief="flat",
    )
    style.map("Treeview", background=[("selected", COLORS["btn_primary"])])
    style.configure(
        "TScrollbar",
        background=COLORS["bg_card"],
        troughcolor=COLORS["bg_main"],
        borderwidth=0,
        relief="flat",
    )
    style.configure(
        "TProgressbar",
        background=COLORS["green"],
        troughcolor=COLORS["bg_card"],
        borderwidth=0,
        relief="flat",
    )
    style.configure(
        "TCombobox",
        background=COLORS["bg_card"],
        foreground=COLORS["text_main"],
        fieldbackground=COLORS["bg_card"],
        arrowcolor=COLORS["text_sub"],
        borderwidth=1,
        relief="flat",
    )


def _make_btn(parent, text: str, command=None, style_type: str = "normal") -> "tk.Button":
    """다크 테마 Button을 생성한다."""
    import tkinter as tk
    bg_map = {
        "normal":  COLORS["btn_bg"],
        "primary": COLORS["btn_primary"],
        "success": COLORS["btn_success"],
        "danger":  "#da3633",
        "warn":    "#9e6a03",
    }
    bg = bg_map.get(style_type, COLORS["btn_bg"])
    fg = "#ffffff" if style_type != "normal" else COLORS["text_head"]
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        bg=bg,
        fg=fg,
        activebackground=COLORS["btn_hover"],
        activeforeground=COLORS["text_main"],
        relief="flat",
        bd=0,
        padx=12,
        pady=5,
        font=FONTS["small"],
        cursor="hand2",
    )
    return btn


def _make_label(parent, text: str, style: str = "normal", **kw) -> "tk.Label":
    """다크 테마 Label을 생성한다."""
    import tkinter as tk
    fg_map = {
        "normal":  COLORS["text_main"],
        "sub":     COLORS["text_sub"],
        "head":    COLORS["text_head"],
        "accent":  COLORS["accent"],
        "green":   COLORS["green"],
        "red":     COLORS["red"],
        "yellow":  COLORS["yellow"],
        "title":   COLORS["accent"],
    }
    font_map = {
        "normal":  FONTS["normal"],
        "sub":     FONTS["small"],
        "head":    FONTS["bold"],
        "accent":  FONTS["normal"],
        "green":   FONTS["normal"],
        "red":     FONTS["normal"],
        "yellow":  FONTS["normal"],
        "title":   FONTS["title"],
    }
    return tk.Label(
        parent,
        text=text,
        fg=fg_map.get(style, COLORS["text_main"]),
        bg=kw.pop("bg", COLORS["bg_main"]),
        font=kw.pop("font", font_map.get(style, FONTS["normal"])),
        **kw,
    )


# ──────────────────────────────────────────────
# 탭 빌더 함수
# ──────────────────────────────────────────────

def _build_topbar(root, app: DocIntelligenceApp) -> dict:
    """상단 바를 구성하고 상태 레이블 참조를 반환한다."""
    import tkinter as tk

    bar = tk.Frame(root, bg=COLORS["bg_panel"], height=48)
    bar.pack(fill="x", side="top")
    bar.pack_propagate(False)

    left = tk.Frame(bar, bg=COLORS["bg_panel"])
    left.pack(side="left", padx=16, pady=0)

    _make_label(left, "Doc Intelligence", style="title", bg=COLORS["bg_panel"]).pack(side="left")
    _make_label(left, "v0.1 MVP", style="sub", bg=COLORS["bg_panel"]).pack(side="left", padx=(6, 0), pady=2)

    right = tk.Frame(bar, bg=COLORS["bg_panel"])
    right.pack(side="right", padx=16)

    # 상태 도트
    dot_canvas = tk.Canvas(right, width=10, height=10, bg=COLORS["bg_panel"], highlightthickness=0)
    dot_canvas.pack(side="left", pady=2)
    dot_canvas.create_oval(1, 1, 9, 9, fill=COLORS["green"], outline="")

    status_lbl = _make_label(right, "COM 연결됨 | 문서 0개 열림", style="sub", bg=COLORS["bg_panel"])
    status_lbl.pack(side="left", padx=(4, 12))

    _make_btn(right, "+ 영역 연결", style_type="primary").pack(side="left", padx=4)
    _make_btn(right, "설정", style_type="normal").pack(side="left", padx=4)

    # 구분선
    sep = tk.Frame(root, bg=COLORS["border"], height=1)
    sep.pack(fill="x")

    return {"status_lbl": status_lbl}


def _build_dashboard_tab(parent, app: DocIntelligenceApp) -> dict:
    """대시보드 탭 UI를 구성한다."""
    import tkinter as tk
    from tkinter import ttk

    container = tk.Frame(parent, bg=COLORS["bg_main"])
    container.pack(fill="both", expand=True)

    # 좌우 분할
    left_frame = tk.Frame(container, bg=COLORS["bg_dark"], width=320)
    left_frame.pack(side="left", fill="y")
    left_frame.pack_propagate(False)

    right_frame = tk.Frame(container, bg=COLORS["bg_main"])
    right_frame.pack(side="left", fill="both", expand=True)

    # ── 좌측: 열린 문서 패널 ──
    doc_header = tk.Frame(left_frame, bg=COLORS["bg_panel"])
    doc_header.pack(fill="x")
    _make_label(doc_header, "열린 문서", style="head", bg=COLORS["bg_panel"]).pack(side="left", padx=12, pady=8)
    doc_count_lbl = _make_label(doc_header, "0개 열림", style="accent", bg=COLORS["bg_panel"])
    doc_count_lbl.pack(side="right", padx=12, pady=8)

    sep1 = tk.Frame(left_frame, bg=COLORS["border"], height=1)
    sep1.pack(fill="x")

    doc_body = tk.Frame(left_frame, bg=COLORS["bg_dark"])
    doc_body.pack(fill="both", expand=True, padx=8, pady=8)

    columns_doc = ("icon", "name", "status")
    doc_tree = ttk.Treeview(doc_body, columns=columns_doc, show="headings", height=20)
    doc_tree.heading("icon", text="유형")
    doc_tree.heading("name", text="파일명")
    doc_tree.heading("status", text="상태")
    doc_tree.column("icon", width=40, anchor="center")
    doc_tree.column("name", width=180)
    doc_tree.column("status", width=60, anchor="center")
    doc_tree.tag_configure("matched", foreground=COLORS["green"])
    doc_tree.tag_configure("new", foreground=COLORS["accent"])
    doc_tree.pack(fill="both", expand=True)

    # ── 우측 상단: 추출 엔티티 패널 ──
    ent_panel = tk.Frame(right_frame, bg=COLORS["bg_dark"])
    ent_panel.pack(fill="both", expand=True, padx=(1, 0), pady=(0, 1))

    ent_header = tk.Frame(ent_panel, bg=COLORS["bg_panel"])
    ent_header.pack(fill="x")
    _make_label(ent_header, "추출된 엔티티", style="head", bg=COLORS["bg_panel"]).pack(side="left", padx=12, pady=8)

    sep2 = tk.Frame(ent_panel, bg=COLORS["border"], height=1)
    sep2.pack(fill="x")

    ent_body = tk.Frame(ent_panel, bg=COLORS["bg_dark"])
    ent_body.pack(fill="both", expand=True, padx=8, pady=8)

    columns_ent = ("type", "value", "confidence")
    ent_tree = ttk.Treeview(ent_body, columns=columns_ent, show="headings", height=8)
    ent_tree.heading("type", text="타입")
    ent_tree.heading("value", text="값")
    ent_tree.heading("confidence", text="신뢰도")
    ent_tree.column("type", width=80)
    ent_tree.column("value", width=200)
    ent_tree.column("confidence", width=70, anchor="center")
    ent_tree.pack(fill="both", expand=True)

    # ── 우측 하단: 활동 로그 패널 ──
    log_panel = tk.Frame(right_frame, bg=COLORS["bg_dark"])
    log_panel.pack(fill="both", expand=True, padx=(1, 0), pady=(1, 0))

    log_header = tk.Frame(log_panel, bg=COLORS["bg_panel"])
    log_header.pack(fill="x")
    _make_label(log_header, "활동 로그", style="head", bg=COLORS["bg_panel"]).pack(side="left", padx=12, pady=8)

    sep3 = tk.Frame(log_panel, bg=COLORS["border"], height=1)
    sep3.pack(fill="x")

    log_body = tk.Frame(log_panel, bg=COLORS["bg_dark"])
    log_body.pack(fill="both", expand=True, padx=8, pady=8)

    columns_log = ("time", "message")
    log_tree = ttk.Treeview(log_body, columns=columns_log, show="headings", height=6)
    log_tree.heading("time", text="시간")
    log_tree.heading("message", text="메시지")
    log_tree.column("time", width=60, anchor="center")
    log_tree.column("message", width=400)
    log_tree.pack(fill="both", expand=True)

    return {
        "doc_tree": doc_tree,
        "ent_tree": ent_tree,
        "log_tree": log_tree,
        "doc_count_lbl": doc_count_lbl,
    }


def _build_learning_tab(parent, app: DocIntelligenceApp) -> dict:
    """학습 모드 탭 UI를 구성한다."""
    import tkinter as tk
    from tkinter import ttk

    container = tk.Frame(parent, bg=COLORS["bg_main"])
    container.pack(fill="both", expand=True)

    # ── 알림 바 ──
    alert_frame = tk.Frame(container, bg=COLORS["bg_blue"], bd=1, relief="flat")
    alert_frame.pack(fill="x", padx=12, pady=12)

    alert_inner = tk.Frame(alert_frame, bg=COLORS["bg_blue"])
    alert_inner.pack(padx=16, pady=10)

    _make_label(alert_inner, "신규 양식 감지", style="accent", bg=COLORS["bg_blue"],
                font=("맑은 고딕", 10, "bold")).pack(side="left")
    _make_label(alert_inner, " — 자동 분석 결과를 확인해 주십시오", style="head",
                bg=COLORS["bg_blue"]).pack(side="left")

    # ── 좌우 분할 ──
    body = tk.Frame(container, bg=COLORS["border"], bd=0)
    body.pack(fill="both", expand=True)

    left_pane = tk.Frame(body, bg=COLORS["bg_dark"])
    left_pane.pack(side="left", fill="both", expand=True)

    right_pane = tk.Frame(body, bg=COLORS["bg_panel"], width=360)
    right_pane.pack(side="left", fill="y")
    right_pane.pack_propagate(False)

    sep_v = tk.Frame(body, bg=COLORS["border"], width=1)
    sep_v.place(relx=0.6, rely=0, relheight=1)

    # ── 좌측: 문서 미리보기 ──
    preview_header = tk.Frame(left_pane, bg=COLORS["bg_panel"])
    preview_header.pack(fill="x")
    _make_label(preview_header, "문서 미리보기", style="sub", bg=COLORS["bg_panel"]).pack(
        side="left", padx=12, pady=6)

    sep_lh = tk.Frame(left_pane, bg=COLORS["border"], height=1)
    sep_lh.pack(fill="x")

    preview_body = tk.Frame(left_pane, bg=COLORS["bg_dark"])
    preview_body.pack(fill="both", expand=True, padx=8, pady=8)

    columns_prev = ("row", "col_a", "col_b", "col_c", "col_d")
    preview_tree = ttk.Treeview(preview_body, columns=columns_prev, show="headings", height=14)
    preview_tree.heading("row", text="#")
    preview_tree.heading("col_a", text="A열")
    preview_tree.heading("col_b", text="B열")
    preview_tree.heading("col_c", text="C열")
    preview_tree.heading("col_d", text="D열")
    preview_tree.column("row", width=30, anchor="center")
    preview_tree.column("col_a", width=100)
    preview_tree.column("col_b", width=100)
    preview_tree.column("col_c", width=100)
    preview_tree.column("col_d", width=100)
    preview_tree.pack(fill="both", expand=True)

    # ── 우측: 자동 분석 결과 편집 ──
    editor_header = tk.Frame(right_pane, bg=COLORS["bg_panel"])
    editor_header.pack(fill="x")
    _make_label(editor_header, "자동 분석 결과", style="head", bg=COLORS["bg_panel"]).pack(
        side="left", padx=12, pady=8)

    sep_rh = tk.Frame(right_pane, bg=COLORS["border"], height=1)
    sep_rh.pack(fill="x")

    editor_body = tk.Frame(right_pane, bg=COLORS["bg_panel"])
    editor_body.pack(fill="both", expand=True, padx=10, pady=8)

    FIELD_TYPES = [
        "날짜", "착공일", "준공일", "검수일", "금액", "예상비용", "부가세",
        "업체명", "부서", "설비코드", "부품코드", "이름", "검수자", "승인자",
        "문서번호", "문서 ID", "무시",
    ]

    sample_fields = [
        ("B3", "2025-03-01", "날짜", "97%"),
        ("C3", "삼성전자", "업체명", "89%"),
        ("D5", "1,500,000", "금액", "95%"),
        ("E7", "정비팀", "부서", "78%"),
    ]

    combo_vars = []
    for loc, val, ftype, conf in sample_fields:
        row_frame = tk.Frame(editor_body, bg=COLORS["bg_card"], bd=0)
        row_frame.pack(fill="x", pady=3)

        _make_label(row_frame, loc, style="sub", bg=COLORS["bg_card"],
                    width=5, anchor="w").pack(side="left", padx=(8, 4), pady=6)
        _make_label(row_frame, val, style="normal", bg=COLORS["bg_card"],
                    width=12, anchor="w").pack(side="left", padx=4)

        var = tk.StringVar(value=ftype)
        combo_vars.append(var)
        combo = ttk.Combobox(row_frame, textvariable=var, values=FIELD_TYPES,
                             state="readonly", width=10)
        combo.pack(side="left", padx=4)

        _make_label(row_frame, conf, style="sub", bg=COLORS["bg_card"]).pack(side="left", padx=4)

    # ── 하단 버튼 ──
    btn_bar = tk.Frame(right_pane, bg=COLORS["bg_panel"])
    btn_bar.pack(fill="x", side="bottom", padx=10, pady=10)

    _make_btn(btn_bar, "템플릿 저장", style_type="success").pack(side="left", padx=4)
    _make_btn(btn_bar, "건너뛰기", style_type="normal").pack(side="left", padx=4)

    return {"preview_tree": preview_tree, "combo_vars": combo_vars}


def _build_rules_tab(parent, app: DocIntelligenceApp) -> dict:
    """룰/프리셋 탭 UI를 구성한다."""
    import tkinter as tk
    from tkinter import ttk

    container = tk.Frame(parent, bg=COLORS["bg_main"])
    container.pack(fill="both", expand=True)

    # 3열 분할: 프리셋 목록 | 룰 목록 | 룰 상세
    preset_pane = tk.Frame(container, bg=COLORS["bg_dark"], width=280)
    preset_pane.pack(side="left", fill="y")
    preset_pane.pack_propagate(False)

    sep1 = tk.Frame(container, bg=COLORS["border"], width=1)
    sep1.pack(side="left", fill="y")

    rule_pane = tk.Frame(container, bg=COLORS["bg_dark"])
    rule_pane.pack(side="left", fill="both", expand=True)

    sep2 = tk.Frame(container, bg=COLORS["border"], width=1)
    sep2.pack(side="left", fill="y")

    detail_pane = tk.Frame(container, bg=COLORS["bg_panel"], width=320)
    detail_pane.pack(side="left", fill="y")
    detail_pane.pack_propagate(False)

    # ── 프리셋 목록 ──
    ph = tk.Frame(preset_pane, bg=COLORS["bg_panel"])
    ph.pack(fill="x")
    _make_label(ph, "프리셋", style="head", bg=COLORS["bg_panel"]).pack(side="left", padx=12, pady=8)
    _make_btn(ph, "+ 추가", style_type="normal").pack(side="right", padx=8, pady=6)

    sep_ph = tk.Frame(preset_pane, bg=COLORS["border"], height=1)
    sep_ph.pack(fill="x")

    preset_lb = tk.Listbox(
        preset_pane,
        bg=COLORS["bg_dark"],
        fg=COLORS["text_main"],
        selectbackground=COLORS["btn_primary"],
        selectforeground="#ffffff",
        font=FONTS["normal"],
        bd=0,
        relief="flat",
        activestyle="none",
    )
    preset_lb.pack(fill="both", expand=True, padx=6, pady=6)

    sample_presets = ["정비비용_기본", "검수 표준 룰셋", "견적서_고급", "발주서_기본"]
    for name in sample_presets:
        preset_lb.insert("end", "  " + name)

    # ── 룰 목록 ──
    rh = tk.Frame(rule_pane, bg=COLORS["bg_panel"])
    rh.pack(fill="x")
    _make_label(rh, "룰 목록", style="head", bg=COLORS["bg_panel"]).pack(side="left", padx=12, pady=8)

    sep_rh = tk.Frame(rule_pane, bg=COLORS["border"], height=1)
    sep_rh.pack(fill="x")

    rule_body = tk.Frame(rule_pane, bg=COLORS["bg_dark"])
    rule_body.pack(fill="both", expand=True, padx=8, pady=8)

    sample_rules = [
        ("금액 일치 검증", "ValueMatch", True),
        ("날짜 순서 검증", "OrderCheck", True),
        ("필수 항목 존재", "Existence", True),
        ("합계 계산 검증", "Calculation", False),
        ("부가세 10% 검증", "Calculation", True),
    ]

    TYPE_COLORS = {
        "ValueMatch":  (COLORS["bg_green"], COLORS["green"]),
        "OrderCheck":  (COLORS["bg_blue"], COLORS["accent"]),
        "Calculation": (COLORS["bg_yellow"], COLORS["yellow"]),
        "Existence":   ("#2d1a3d", "#bc8cff"),
    }

    for rname, rtype, active in sample_rules:
        card = tk.Frame(rule_body, bg=COLORS["bg_panel"], bd=0)
        card.pack(fill="x", pady=3)
        card.configure(highlightbackground=COLORS["border2"], highlightthickness=1)

        chk_bg = COLORS["btn_success"] if active else COLORS["bg_card"]
        chk_lbl = tk.Label(card, text="✓" if active else " ", bg=chk_bg,
                           fg="#ffffff", width=2, font=FONTS["small"])
        chk_lbl.pack(side="left", padx=(10, 8), pady=8)

        info_frame = tk.Frame(card, bg=COLORS["bg_panel"])
        info_frame.pack(side="left", fill="x", expand=True)
        _make_label(info_frame, rname, style="normal", bg=COLORS["bg_panel"]).pack(anchor="w")

        tbg, tfg = TYPE_COLORS.get(rtype, (COLORS["bg_card"], COLORS["text_sub"]))
        type_lbl = tk.Label(card, text=rtype, bg=tbg, fg=tfg,
                            font=FONTS["small"], padx=6, pady=2)
        type_lbl.pack(side="right", padx=10)

    # ── 룰 상세 ──
    dh = tk.Frame(detail_pane, bg=COLORS["bg_panel"])
    dh.pack(fill="x")
    _make_label(dh, "룰 상세", style="head", bg=COLORS["bg_panel"]).pack(side="left", padx=12, pady=8)

    sep_dh = tk.Frame(detail_pane, bg=COLORS["border"], height=1)
    sep_dh.pack(fill="x")

    detail_body = tk.Frame(detail_pane, bg=COLORS["bg_panel"])
    detail_body.pack(fill="both", expand=True, padx=12, pady=12)

    _make_label(detail_body, "연결된 영역", style="sub", bg=COLORS["bg_panel"]).pack(anchor="w", pady=(0, 6))

    for doc_name, loc in [("정산서_2025.xlsx", "B3:D10"), ("견적서.docx", "§금액 섹션")]:
        region_card = tk.Frame(detail_body, bg=COLORS["bg_card"], bd=0)
        region_card.pack(fill="x", pady=3)
        _make_label(region_card, doc_name, style="accent", bg=COLORS["bg_card"]).pack(
            anchor="w", padx=10, pady=(6, 2))
        _make_label(region_card, loc, style="sub", bg=COLORS["bg_card"]).pack(
            anchor="w", padx=10, pady=(0, 6))

    return {"preset_lb": preset_lb}


def _build_linker_tab(parent, app: DocIntelligenceApp) -> dict:
    """영역 연결 탭 UI를 구성한다."""
    import tkinter as tk

    container = tk.Frame(parent, bg=COLORS["bg_main"])
    container.pack(fill="both", expand=True)

    # ── 안내 영역 ──
    guide_frame = tk.Frame(container, bg=COLORS["bg_panel"])
    guide_frame.pack(fill="x")

    _make_label(guide_frame, "영역 연결 모드를 시작하려면 버튼을 클릭하십시오.",
                style="sub", bg=COLORS["bg_panel"]).pack(side="left", padx=16, pady=12)

    def _start_overlay():
        try:
            from doc_intelligence.ui_components import OverlayWindow
            def _on_region(x1, y1, x2, y2):
                msg = f"영역 선택 완료: ({x1},{y1})→({x2},{y2})"
                app.add_activity_log(msg)
                region_list.insert("end", f"  [{x2-x1}×{y2-y1}] {msg}")
            OverlayWindow(on_region_selected=_on_region)
        except Exception as e:
            logger.warning("OverlayWindow 오류: %s", e)

    _make_btn(guide_frame, "영역 연결 시작", command=_start_overlay,
              style_type="primary").pack(side="right", padx=16, pady=8)

    sep_g = tk.Frame(container, bg=COLORS["border"], height=1)
    sep_g.pack(fill="x")

    # ── 단계 표시 ──
    steps_frame = tk.Frame(container, bg=COLORS["bg_panel"])
    steps_frame.pack(fill="x", padx=16, pady=8)

    steps = [
        ("1", "기준 문서 선택", True, False),
        ("2", "영역 드래그", False, False),
        ("3", "대상 문서 선택", False, False),
        ("4", "영역 드래그", False, False),
        ("5", "연결 확인", False, False),
    ]

    for num, label, is_active, is_done in steps:
        step_bg = COLORS["btn_primary"] if is_active else (COLORS["btn_success"] if is_done else COLORS["bg_card"])
        step_fg = "#ffffff"
        num_lbl = tk.Label(steps_frame, text=num, bg=step_bg, fg=step_fg,
                           font=("맑은 고딕", 10, "bold"), width=2, height=1)
        num_lbl.pack(side="left", padx=(0, 4))
        txt_fg = COLORS["text_head"] if is_active else COLORS["text_sub"]
        _make_label(steps_frame, label, style="normal", bg=COLORS["bg_panel"],
                    fg=txt_fg).pack(side="left")
        if num != "5":
            _make_label(steps_frame, "→", style="sub", bg=COLORS["bg_panel"]).pack(side="left", padx=6)

    sep_s = tk.Frame(container, bg=COLORS["border"], height=1)
    sep_s.pack(fill="x")

    # ── 연결된 영역 목록 ──
    list_header = tk.Frame(container, bg=COLORS["bg_panel"])
    list_header.pack(fill="x")
    _make_label(list_header, "연결된 영역 목록", style="head", bg=COLORS["bg_panel"]).pack(
        side="left", padx=12, pady=8)

    sep_lh = tk.Frame(container, bg=COLORS["border"], height=1)
    sep_lh.pack(fill="x")

    list_body = tk.Frame(container, bg=COLORS["bg_dark"])
    list_body.pack(fill="both", expand=True, padx=8, pady=8)

    region_list = tk.Listbox(
        list_body,
        bg=COLORS["bg_dark"],
        fg=COLORS["text_main"],
        selectbackground=COLORS["btn_primary"],
        font=FONTS["normal"],
        bd=0,
        relief="flat",
        activestyle="none",
    )
    region_list.pack(fill="both", expand=True)

    return {"region_list": region_list}


def _build_validation_tab(parent, app: DocIntelligenceApp) -> dict:
    """검증 결과 탭 UI를 구성한다."""
    import tkinter as tk
    from tkinter import ttk

    container = tk.Frame(parent, bg=COLORS["bg_main"])
    container.pack(fill="both", expand=True)

    # 좌우 분할
    result_pane = tk.Frame(container, bg=COLORS["bg_dark"])
    result_pane.pack(side="left", fill="both", expand=True)

    sep = tk.Frame(container, bg=COLORS["border"], width=1)
    sep.pack(side="left", fill="y")

    summary_pane = tk.Frame(container, bg=COLORS["bg_panel"], width=340)
    summary_pane.pack(side="left", fill="y")
    summary_pane.pack_propagate(False)

    # ── 좌측: 결과 리스트 ──
    rh = tk.Frame(result_pane, bg=COLORS["bg_panel"])
    rh.pack(fill="x")
    _make_label(rh, "검증 결과", style="head", bg=COLORS["bg_panel"]).pack(side="left", padx=12, pady=8)

    sep_rh = tk.Frame(result_pane, bg=COLORS["border"], height=1)
    sep_rh.pack(fill="x")

    result_body = tk.Frame(result_pane, bg=COLORS["bg_dark"])
    result_body.pack(fill="both", expand=True, padx=8, pady=8)

    columns_res = ("icon", "rule", "status", "detail")
    result_tree = ttk.Treeview(result_body, columns=columns_res, show="headings", height=18)
    result_tree.heading("icon", text="")
    result_tree.heading("rule", text="룰 이름")
    result_tree.heading("status", text="상태")
    result_tree.heading("detail", text="상세")
    result_tree.column("icon", width=24, anchor="center")
    result_tree.column("rule", width=140)
    result_tree.column("status", width=70, anchor="center")
    result_tree.column("detail", width=300)
    result_tree.tag_configure("pass", foreground=COLORS["green"])
    result_tree.tag_configure("fail", foreground=COLORS["red"])
    result_tree.tag_configure("warn", foreground=COLORS["yellow"])
    result_tree.pack(fill="both", expand=True)

    sample_results = [
        ("✓", "금액 일치 검증", "통과", "정산서 B10 = 견적서 D5 (1,500,000)", "pass"),
        ("✗", "날짜 순서 검증", "실패", "착공일(03.15) > 준공일(03.10) — 역순 감지", "fail"),
        ("!", "부가세 검증", "경고", "부가세 149,900 ≠ 금액×10% (150,000)", "warn"),
        ("✓", "필수항목 존재", "통과", "업체명, 설비코드, 검수자 모두 존재", "pass"),
        ("✓", "합계 계산 검증", "통과", "소계 합산 일치", "pass"),
    ]

    for icon, rule, status, detail, tag in sample_results:
        result_tree.insert("", "end", values=(icon, rule, status, detail), tags=(tag,))

    # ── 우측: 요약 패널 ──
    sh = tk.Frame(summary_pane, bg=COLORS["bg_panel"])
    sh.pack(fill="x")
    _make_label(sh, "검증 요약", style="head", bg=COLORS["bg_panel"]).pack(side="left", padx=12, pady=8)

    sep_sh = tk.Frame(summary_pane, bg=COLORS["border"], height=1)
    sep_sh.pack(fill="x")

    summary_body = tk.Frame(summary_pane, bg=COLORS["bg_panel"])
    summary_body.pack(fill="both", expand=True, padx=16, pady=16)

    # 통과/실패/경고 카운트
    stat_card = tk.Frame(summary_body, bg=COLORS["bg_card"], bd=0)
    stat_card.pack(fill="x", pady=(0, 12))
    stat_card.configure(highlightbackground=COLORS["border2"], highlightthickness=1)

    stat_inner = tk.Frame(stat_card, bg=COLORS["bg_card"])
    stat_inner.pack(padx=16, pady=12)

    _make_label(stat_inner, "검증 결과 요약", style="sub", bg=COLORS["bg_card"]).pack(anchor="w", pady=(0, 8))

    counts_frame = tk.Frame(stat_inner, bg=COLORS["bg_card"])
    counts_frame.pack(fill="x")

    pass_count_lbl = tk.Label(counts_frame, text="3", fg=COLORS["green"],
                              bg=COLORS["bg_card"], font=("맑은 고딕", 22, "bold"))
    pass_count_lbl.grid(row=0, column=0, padx=10)
    tk.Label(counts_frame, text="통과", fg=COLORS["text_sub"],
             bg=COLORS["bg_card"], font=FONTS["small"]).grid(row=1, column=0)

    fail_count_lbl = tk.Label(counts_frame, text="1", fg=COLORS["red"],
                              bg=COLORS["bg_card"], font=("맑은 고딕", 22, "bold"))
    fail_count_lbl.grid(row=0, column=1, padx=10)
    tk.Label(counts_frame, text="실패", fg=COLORS["text_sub"],
             bg=COLORS["bg_card"], font=FONTS["small"]).grid(row=1, column=1)

    warn_count_lbl = tk.Label(counts_frame, text="1", fg=COLORS["yellow"],
                              bg=COLORS["bg_card"], font=("맑은 고딕", 22, "bold"))
    warn_count_lbl.grid(row=0, column=2, padx=10)
    tk.Label(counts_frame, text="경고", fg=COLORS["text_sub"],
             bg=COLORS["bg_card"], font=FONTS["small"]).grid(row=1, column=2)

    # 프로그레스바
    _make_label(summary_body, "통과율", style="sub", bg=COLORS["bg_panel"]).pack(anchor="w", pady=(8, 4))
    progress = ttk.Progressbar(summary_body, value=60, maximum=100,
                               mode="determinate", length=280)
    progress.pack(fill="x")

    progress_lbl = _make_label(summary_body, "60%", style="accent", bg=COLORS["bg_panel"])
    progress_lbl.pack(anchor="e", pady=(2, 0))

    return {
        "result_tree": result_tree,
        "pass_count_lbl": pass_count_lbl,
        "fail_count_lbl": fail_count_lbl,
        "warn_count_lbl": warn_count_lbl,
        "progress": progress,
        "progress_lbl": progress_lbl,
    }


# ──────────────────────────────────────────────
# UI 갱신 함수
# ──────────────────────────────────────────────

def _update_ui(root, app: DocIntelligenceApp, ui_refs: dict) -> None:
    """UI 큐를 처리하고 위젯을 갱신한다. root.after로 반복 호출."""
    # 큐 처리
    while not app._ui_queue.empty():
        try:
            event = app._ui_queue.get_nowait()
            if event[0] == "doc_processed":
                _, doc_info, context = event
                name = doc_info.get("name", "알 수 없음")
                app.open_docs.append(doc_info)
                app.entities.extend(context.get("entities", []))
                app.add_activity_log(f"문서 처리 완료: {name}")
        except queue.Empty:
            break

    # 문서 목록 갱신
    doc_tree = ui_refs.get("doc_tree")
    if doc_tree:
        for item in doc_tree.get_children():
            doc_tree.delete(item)
        for doc in app.open_docs:
            app_key = _APP_TO_PARSER.get(doc.get("app", ""), "")
            icon = _APP_TO_ICON.get(app_key, "?")
            name = doc.get("name", "")
            doc_tree.insert("", "end", values=(icon, name, "감지됨"), tags=("new",))

    # 문서 카운트 갱신
    doc_count_lbl = ui_refs.get("doc_count_lbl")
    if doc_count_lbl:
        doc_count_lbl.config(text=f"{len(app.open_docs)}개 열림")

    # 상단 상태 레이블 갱신
    status_lbl = ui_refs.get("status_lbl")
    if status_lbl:
        status_lbl.config(text=f"COM 연결됨 | 문서 {len(app.open_docs)}개 열림")

    # 엔티티 목록 갱신
    ent_tree = ui_refs.get("ent_tree")
    if ent_tree:
        for item in ent_tree.get_children():
            ent_tree.delete(item)
        for entity in app.entities[-50:]:
            if hasattr(entity, "type"):
                ent_tree.insert("", "end", values=(
                    entity.type, entity.value, f"{entity.confidence:.0%}"))
            elif isinstance(entity, dict):
                ent_tree.insert("", "end", values=(
                    entity.get("type", ""),
                    entity.get("value", ""),
                    entity.get("confidence", ""),
                ))

    # 활동 로그 갱신
    log_tree = ui_refs.get("log_tree")
    if log_tree:
        current_count = len(log_tree.get_children())
        log_entries = app.activity_log
        if len(log_entries) > current_count:
            for entry in log_entries[current_count:]:
                log_tree.insert("", "end", values=(entry["time"], entry["message"]))
            # 최신 항목으로 스크롤
            children = log_tree.get_children()
            if children:
                log_tree.see(children[-1])

    # 1초 후 재호출
    root.after(1000, lambda: _update_ui(root, app, ui_refs))


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    )

    app = DocIntelligenceApp()

    try:
        import tkinter as tk
        from tkinter import ttk

        root = tk.Tk()
        root.title("Doc Intelligence v0.1")
        root.geometry("1280x800")
        root.configure(bg=COLORS["bg_main"])
        root.minsize(900, 600)

        style = ttk.Style()
        _apply_dark_style(style)

        # ── 상단 바 ──
        topbar_refs = _build_topbar(root, app)

        # ── 탭 바 (ttk.Notebook) ──
        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True)

        tab_dashboard  = tk.Frame(notebook, bg=COLORS["bg_main"])
        tab_learning   = tk.Frame(notebook, bg=COLORS["bg_main"])
        tab_rules      = tk.Frame(notebook, bg=COLORS["bg_main"])
        tab_linker     = tk.Frame(notebook, bg=COLORS["bg_main"])
        tab_validation = tk.Frame(notebook, bg=COLORS["bg_main"])

        notebook.add(tab_dashboard,  text="대시보드")
        notebook.add(tab_learning,   text="학습 모드")
        notebook.add(tab_rules,      text="룰 / 프리셋")
        notebook.add(tab_linker,     text="영역 연결")
        notebook.add(tab_validation, text="검증 결과")

        # ── 각 탭 빌드 ──
        dash_refs   = _build_dashboard_tab(tab_dashboard, app)
        learn_refs  = _build_learning_tab(tab_learning, app)
        rules_refs  = _build_rules_tab(tab_rules, app)
        linker_refs = _build_linker_tab(tab_linker, app)
        valid_refs  = _build_validation_tab(tab_validation, app)

        # ── UI 참조 통합 ──
        ui_refs = {
            **topbar_refs,
            **dash_refs,
        }

        # ── 초기 활동 로그 ──
        app.add_activity_log("Doc Intelligence 시작됨")
        app.add_activity_log("COM 폴링 초기화 중...")

        # ── COM 폴링 시작 ──
        app.start_polling()

        # ── UI 갱신 루프 시작 ──
        root.after(1000, lambda: _update_ui(root, app, ui_refs))

        root.mainloop()

    except Exception as exc:
        logger.exception("tkinter GUI 시작 실패: %s", exc)
        app.start_polling()
        input("Enter를 누르면 종료합니다...")
    finally:
        app.stop_polling()


if __name__ == "__main__":
    main()
