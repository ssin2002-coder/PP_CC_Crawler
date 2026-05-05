import { useEffect, useRef, useCallback } from 'react';
import { io } from 'socket.io-client';
import useStore from '../stores/store.js';
import { fetchData, runValidation } from '../utils/api.js';

// =============================================
// Socket.IO WebSocket 훅
// 서버와의 실시간 연결 + Excel 자동 인식
// =============================================

const RECONNECT_DELAYS = [1000, 2000, 4000, 8000, 16000];

export function useWebSocket() {
  const socketRef = useRef(null);
  const reconnectAttemptRef = useRef(0);
  const reconnectTimerRef = useRef(null);

  const setConnected = useStore((s) => s.setConnected);
  const setExcelData = useStore((s) => s.setExcelData);
  const setActiveWorkbook = useStore((s) => s.setActiveWorkbook);
  const setActiveSheet = useStore((s) => s.setActiveSheet);
  const setIssues = useStore((s) => s.setIssues);
  const setSummary = useStore((s) => s.setSummary);

  // 워크북 감지 시 자동 선택 + 데이터 로드
  const handleAutoSelect = useCallback(async (workbooks) => {
    if (!workbooks || workbooks.length === 0) {
      setExcelData({ workbooks: [], activeWorkbook: null, activeSheet: null, sheets: [], cells: [] });
      setConnected(false);
      return;
    }

    const store = useStore.getState();
    const currentWb = store.excelData.activeWorkbook;
    const currentWbName = currentWb ? (typeof currentWb === 'string' ? currentWb : currentWb.name) : null;

    // 이미 같은 워크북이 선택되어 있으면 워크북 목록만 갱신
    const alreadySelected = currentWbName && workbooks.some((wb) => wb.name === currentWbName);

    setExcelData({ workbooks });
    setConnected(true);

    if (!alreadySelected) {
      // 첫 번째 워크북 자동 선택
      const first = workbooks[0];
      const firstSheet = first.sheets && first.sheets.length > 0 ? first.sheets[0] : null;

      setExcelData({
        activeWorkbook: first,
        sheets: first.sheets || [],
        activeSheet: firstSheet,
      });

      // 셀 데이터 자동 로드 + 자동 검증
      if (firstSheet) {
        const { data } = await fetchData(first.name, firstSheet);
        if (data && data.cells) {
          setExcelData({ cells: data.cells });
        }
        // 자동 검증 실행
        const { data: valData } = await runValidation(first.name, firstSheet);
        if (valData) {
          if (valData.issues) setIssues(valData.issues);
          if (valData.summary) setSummary(valData.summary);
        }
      }
      console.log('[WS] 워크북 자동 선택 + 검증:', first.name);
    }
  }, [setExcelData, setConnected]);

  const connect = useCallback(() => {
    if (socketRef.current) {
      socketRef.current.removeAllListeners();
      socketRef.current.disconnect();
    }

    const socket = io('/', {
      path: '/socket.io',
      transports: ['websocket', 'polling'],
      reconnection: false,
    });

    socketRef.current = socket;

    // 연결
    socket.on('connect', () => {
      console.log('[WS] 서버 연결됨:', socket.id);
      setConnected(true);
      reconnectAttemptRef.current = 0;
    });

    socket.on('disconnect', (reason) => {
      console.log('[WS] 서버 연결 해제:', reason);
      setConnected(false);
      scheduleReconnect();
    });

    socket.on('connect_error', (err) => {
      console.warn('[WS] 연결 오류:', err.message);
      setConnected(false);
      scheduleReconnect();
    });

    // --------------------------------------------------
    // Excel 상태 이벤트 (백엔드가 connect 시 + 폴링 시 전송)
    // --------------------------------------------------
    socket.on('excel:status', (data) => {
      // {connected: bool, workbooks: [{name, path, sheets}, ...]}
      handleAutoSelect(data.workbooks || []);
    });

    // Excel 데이터 변경 (폴링에서 감지)
    socket.on('excel:data_changed', async (data) => {
      if (data.cells) {
        setExcelData({ cells: data.cells });
      }
      if (data.workbook) {
        const store = useStore.getState();
        const currentWb = store.excelData.activeWorkbook;
        const currentName = currentWb ? (typeof currentWb === 'string' ? currentWb : currentWb.name) : null;
        if (data.workbook !== currentName) {
          // 새 워크북이면 자동 선택
          handleAutoSelect([{ name: data.workbook, sheets: [data.sheet] }]);
        }
      }
      if (data.sheet) {
        setExcelData({ activeSheet: data.sheet });
      }
    });

    // 검증 결과 (WebSocket 경유 검증 시)
    socket.on('validation:result', (data) => {
      if (data.issues) setIssues(data.issues);
      if (data.summary) setSummary(data.summary);
    });
  }, [setConnected, setExcelData, setIssues, setSummary, handleAutoSelect]);

  const scheduleReconnect = useCallback(() => {
    if (reconnectTimerRef.current) clearTimeout(reconnectTimerRef.current);
    const attempt = reconnectAttemptRef.current;
    const delay = RECONNECT_DELAYS[Math.min(attempt, RECONNECT_DELAYS.length - 1)];
    console.log(`[WS] ${delay}ms 후 재연결 시도 (${attempt + 1}회차)`);
    reconnectTimerRef.current = setTimeout(() => {
      reconnectAttemptRef.current += 1;
      connect();
    }, delay);
  }, [connect]);

  useEffect(() => {
    connect();
    return () => {
      if (reconnectTimerRef.current) clearTimeout(reconnectTimerRef.current);
      if (socketRef.current) {
        socketRef.current.removeAllListeners();
        socketRef.current.disconnect();
      }
    };
  }, [connect]);

  const emit = useCallback((event, data) => {
    if (socketRef.current && socketRef.current.connected) {
      socketRef.current.emit(event, data);
    }
  }, []);

  const isConnected = useStore((s) => s.isConnected);
  return { isConnected, emit };
}
