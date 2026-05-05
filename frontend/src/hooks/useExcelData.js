import { useEffect, useState, useCallback } from 'react';
import useStore from '../stores/store.js';
import { fetchWorkbooks, fetchSheets, fetchData, runValidation } from '../utils/api.js';

// =============================================
// Excel 데이터 훅
// 초기 로드 + WebSocket 자동 갱신 보조
// =============================================

export function useExcelData() {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const excelData = useStore((s) => s.excelData);
  const setExcelData = useStore((s) => s.setExcelData);
  const setActiveSheet = useStore((s) => s.setActiveSheet);
  const setIssues = useStore((s) => s.setIssues);
  const setSummary = useStore((s) => s.setSummary);

  // 워크북 목록 조회 (마운트 시 1회 — WebSocket 연결 전 폴백)
  const loadWorkbooks = useCallback(async () => {
    setLoading(true);
    setError(null);
    const { data, error: err } = await fetchWorkbooks();
    if (err) {
      setError(err);
    } else if (data) {
      const workbooks = data.workbooks || [];
      setExcelData({ workbooks });

      // WebSocket이 아직 워크북을 선택하지 않았으면 자동 선택
      const store = useStore.getState();
      if (!store.excelData.activeWorkbook && workbooks.length > 0) {
        const first = workbooks[0];
        const firstSheet = first.sheets && first.sheets.length > 0 ? first.sheets[0] : null;
        setExcelData({
          activeWorkbook: first,
          sheets: first.sheets || [],
          activeSheet: firstSheet,
        });

        if (firstSheet) {
          const { data: cellData } = await fetchData(first.name, firstSheet);
          if (cellData && cellData.cells) {
            setExcelData({ cells: cellData.cells });
          }
          // 자동 검증
          const { data: valData } = await runValidation(first.name, firstSheet);
          if (valData) {
            if (valData.issues) setIssues(valData.issues);
            if (valData.summary) setSummary(valData.summary);
          }
        }
      }
    }
    setLoading(false);
  }, [setExcelData]);

  // 시트 전환 시 데이터 로드
  const switchSheet = useCallback(async (sheetName) => {
    const wb = excelData.activeWorkbook;
    if (!wb || !sheetName) return;
    const wbName = typeof wb === 'string' ? wb : wb.name;

    setActiveSheet(sheetName);
    setLoading(true);
    const { data, error: err } = await fetchData(wbName, sheetName);
    if (err) {
      setError(err);
    } else if (data) {
      setExcelData({ cells: data.cells || [] });
    }
    // 시트 전환 시 자동 검증
    const { data: valData } = await runValidation(wbName, sheetName);
    if (valData) {
      if (valData.issues) setIssues(valData.issues);
      if (valData.summary) setSummary(valData.summary);
    }
    setLoading(false);
  }, [excelData.activeWorkbook, setExcelData, setActiveSheet]);

  // 마운트 시 초기 로드
  useEffect(() => {
    loadWorkbooks();
  }, [loadWorkbooks]);

  return {
    workbooks: excelData.workbooks,
    sheets: excelData.sheets,
    cells: excelData.cells,
    activeWorkbook: excelData.activeWorkbook,
    activeSheet: excelData.activeSheet,
    loading,
    error,
    refresh: loadWorkbooks,
    switchSheet,
  };
}
