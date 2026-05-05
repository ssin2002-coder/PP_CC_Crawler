import { useState, useCallback } from 'react';
import useStore from '../stores/store.js';
import { runValidation as apiRunValidation, exportResult as apiExportResult } from '../utils/api.js';

// =============================================
// 검증 관련 훅
// 검증 실행 및 결과 내보내기
// =============================================

export function useValidation() {
  const [isValidating, setIsValidating] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  const issues = useStore((s) => s.issues);
  const summary = useStore((s) => s.summary);
  const setIssues = useStore((s) => s.setIssues);
  const setSummary = useStore((s) => s.setSummary);
  const excelData = useStore((s) => s.excelData);

  // 검증 실행
  const runValidation = useCallback(async () => {
    const { activeWorkbook, activeSheet } = excelData;
    if (!activeWorkbook) {
      console.warn('[검증] 활성 워크북 없음');
      return { error: '활성 워크북이 없습니다.' };
    }
    const wbName = typeof activeWorkbook === 'string' ? activeWorkbook : activeWorkbook.name;

    setIsValidating(true);
    const { data, error } = await apiRunValidation(wbName, activeSheet);
    setIsValidating(false);

    if (error) {
      console.error('[검증] 실행 실패:', error);
      return { error };
    }

    if (data) {
      setIssues(data.issues || []);
      setSummary(data.summary || { errors: 0, warnings: 0, info: 0, ok: 0 });
    }

    return { data };
  }, [excelData, setIssues, setSummary]);

  // 결과 내보내기
  const exportResult = useCallback(async () => {
    const { activeWorkbook } = excelData;
    if (!activeWorkbook) {
      console.warn('[내보내기] 활성 워크북 없음');
      return { error: '활성 워크북이 없습니다.' };
    }
    const wbName = typeof activeWorkbook === 'string' ? activeWorkbook : activeWorkbook.name;
    const resultPayload = { issues, summary };

    setIsExporting(true);
    const { data, error } = await apiExportResult(wbName, resultPayload);
    setIsExporting(false);

    if (error) {
      console.error('[내보내기] 실패:', error);
      return { error };
    }

    return { data };
  }, [excelData]);

  return {
    issues,
    summary,
    isValidating,
    isExporting,
    runValidation,
    exportResult,
  };
}
