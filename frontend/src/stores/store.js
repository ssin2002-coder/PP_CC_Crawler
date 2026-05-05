import { create } from 'zustand';

// =============================================
// 메인 Zustand 스토어
// Excel 데이터, 규칙, 검증 결과, UI 상태를 통합 관리
// =============================================

const useStore = create((set, get) => ({
  // --------------------------------------------------
  // Excel 데이터 슬라이스
  // --------------------------------------------------
  excelData: {
    workbooks: [],      // 연결된 워크북 목록
    activeWorkbook: null, // 현재 선택된 워크북 경로
    activeSheet: null,    // 현재 선택된 시트명
    sheets: [],           // 현재 워크북의 시트 목록
    cells: [],            // 현재 시트의 셀 데이터 배열 [{row, col, value, ...}]
  },

  setExcelData: (data) =>
    set((state) => ({
      excelData: { ...state.excelData, ...data },
    })),

  setActiveWorkbook: (workbook) =>
    set((state) => ({
      excelData: {
        ...state.excelData,
        activeWorkbook: workbook,
        activeSheet: null,
        sheets: [],
        cells: [],
      },
    })),

  setActiveSheet: (sheet) =>
    set((state) => ({
      excelData: {
        ...state.excelData,
        activeSheet: sheet,
        cells: [],
      },
    })),

  // --------------------------------------------------
  // 규칙 슬라이스
  // --------------------------------------------------
  rules: [], // [{id, name, description, type, severity, enabled, blocks, issueCount}]

  setRules: (rules) => set({ rules }),

  addRule: (rule) =>
    set((state) => ({ rules: [...state.rules, rule] })),

  updateRule: (id, updates) =>
    set((state) => ({
      rules: state.rules.map((r) => (r.id === id ? { ...r, ...updates } : r)),
    })),

  deleteRule: (id) =>
    set((state) => ({
      rules: state.rules.filter((r) => r.id !== id),
    })),

  toggleRule: (id) =>
    set((state) => ({
      rules: state.rules.map((r) =>
        r.id === id ? { ...r, enabled: !r.enabled } : r
      ),
    })),

  // --------------------------------------------------
  // 검증 결과 슬라이스
  // --------------------------------------------------
  issues: [], // [{id, cellRef, row, col, ruleId, ruleName, severity, message, currentValue, expectedValue}]
  summary: { errors: 0, warnings: 0, info: 0, ok: 0 },

  setIssues: (issues) => set({ issues }),

  setSummary: (summary) => set({ summary }),

  // --------------------------------------------------
  // 선택 상태 슬라이스 (3-way sync: 셀 <-> 규칙 <-> 이슈)
  // --------------------------------------------------
  selectedCell: null,  // {row, col, ref} 형태
  activeRule: null,    // 규칙 id
  selectedIssue: null, // 이슈 id

  setSelectedCell: (cell) => set({ selectedCell: cell }),

  setActiveRule: (ruleId) => set({ activeRule: ruleId }),

  setSelectedIssue: (issueId) => {
    const { issues } = get();
    const issue = issues.find((i) => i.id === issueId);
    const updates = { selectedIssue: issueId };
    if (issue) {
      // 이슈 선택 시 관련 셀과 규칙도 자동 동기화 (snake_case 대응)
      const ref = issue.cell_ref || issue.cellRef || '';
      updates.selectedCell = { row: issue.row, col: issue.col, ref };
      updates.activeRule = issue.rule_id || issue.ruleId;
    }
    set(updates);
  },

  // --------------------------------------------------
  // UI 상태 슬라이스
  // --------------------------------------------------
  isRuleEditorOpen: false,
  editingRule: null, // 편집 중인 규칙 객체 (null이면 신규 생성)

  setRuleEditorOpen: (open) => set({ isRuleEditorOpen: open }),

  setEditingRule: (rule) =>
    set({ editingRule: rule, isRuleEditorOpen: true }),

  closeRuleEditor: () =>
    set({ isRuleEditorOpen: false, editingRule: null }),

  // --------------------------------------------------
  // 연결 상태 슬라이스
  // --------------------------------------------------
  isConnected: false,

  setConnected: (connected) => set({ isConnected: connected }),

  // --------------------------------------------------
  // 계산 헬퍼 함수
  // --------------------------------------------------

  // 특정 셀 참조(예: "E9")에 해당하는 이슈 목록 반환
  getRelatedIssues: (cellRef) => {
    const { issues } = get();
    if (!cellRef) return [];
    return issues.filter((issue) => (issue.cell_ref || issue.cellRef) === cellRef);
  },

  // 특정 규칙 ID에 영향받는 셀 참조 목록 반환
  getRelatedCells: (ruleId) => {
    const { issues } = get();
    if (!ruleId) return [];
    return issues
      .filter((issue) => (issue.rule_id || issue.ruleId) === ruleId)
      .map((issue) => issue.cell_ref || issue.cellRef);
  },
}));

export default useStore;
