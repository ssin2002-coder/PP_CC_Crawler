// =============================================
// API 유틸리티
// Flask 백엔드와의 모든 HTTP 통신을 담당
// =============================================

const BASE_URL = '';

/**
 * 공통 fetch 래퍼 - 에러 처리 포함
 * @param {string} path - API 경로
 * @param {RequestInit} options - fetch 옵션
 * @returns {Promise<{data: any, error: string|null}>}
 */
async function request(path, options = {}) {
  try {
    const res = await fetch(`${BASE_URL}${path}`, {
      headers: { 'Content-Type': 'application/json', ...options.headers },
      ...options,
    });
    if (!res.ok) {
      const text = await res.text();
      return { data: null, error: `HTTP ${res.status}: ${text}` };
    }
    const data = await res.json();
    return { data, error: null };
  } catch (err) {
    return { data: null, error: err.message };
  }
}

// --------------------------------------------------
// Excel 데이터 관련 API
// --------------------------------------------------

/** 연결된 워크북 목록 조회 */
export async function fetchWorkbooks() {
  return request('/api/excel/workbooks');
}

/** 특정 워크북의 시트 목록 조회 */
export async function fetchSheets(workbookName) {
  return request(`/api/excel/sheets?workbook=${encodeURIComponent(workbookName)}`);
}

/** 특정 시트의 셀 데이터 조회 */
export async function fetchData(workbookName, sheetName) {
  return request(
    `/api/excel/data?workbook=${encodeURIComponent(workbookName)}&sheet=${encodeURIComponent(sheetName)}`
  );
}

/** Excel에서 특정 셀로 이동 */
export async function navigateToCell(workbookName, sheetName, cellRef) {
  return request('/api/excel/navigate', {
    method: 'POST',
    body: JSON.stringify({ workbook: workbookName, sheet: sheetName, cell: cellRef }),
  });
}

/** Excel 연결 상태 조회 */
export async function fetchStatus() {
  return request('/api/excel/status');
}

// --------------------------------------------------
// 규칙 관련 API
// --------------------------------------------------

/** 규칙 목록 전체 조회 */
export async function fetchRules() {
  return request('/api/rules');
}

/** 새 규칙 생성 */
export async function createRule(rule) {
  return request('/api/rules', {
    method: 'POST',
    body: JSON.stringify(rule),
  });
}

/** 기존 규칙 수정 */
export async function updateRule(id, updates) {
  return request(`/api/rules/${id}`, {
    method: 'PUT',
    body: JSON.stringify(updates),
  });
}

/** 규칙 삭제 */
export async function deleteRule(id) {
  return request(`/api/rules/${id}`, { method: 'DELETE' });
}

/** 규칙 활성화/비활성화 토글 */
export async function toggleRule(id) {
  return request(`/api/rules/${id}/toggle`, { method: 'PATCH' });
}

// --------------------------------------------------
// 검증 관련 API
// --------------------------------------------------

/** 검증 실행 */
export async function runValidation(workbookName, sheetName) {
  return request('/api/validate/run', {
    method: 'POST',
    body: JSON.stringify({ workbook: workbookName, sheet: sheetName }),
  });
}

/** 검증 결과 내보내기 (JSON 파일 저장) */
export async function exportResult(workbookName, result) {
  return request('/api/validate/export', {
    method: 'POST',
    body: JSON.stringify({ workbook: workbookName, result }),
  });
}

/** 저장된 검증 결과 목록 조회 */
export async function fetchResults() {
  return request('/api/validate/results');
}
