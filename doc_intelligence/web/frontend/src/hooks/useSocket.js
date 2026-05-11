import { useEffect } from 'react';
import { io } from 'socket.io-client';
import { useStore } from '../stores/store';
let socket = null;
export function useSocket() {
  const setDocuments = useStore((s) => s.setDocuments);
  const setComStatus = useStore((s) => s.setComStatus);
  const setEnvStatus = useStore((s) => s.setEnvStatus);
  useEffect(() => {
    socket = io({ transports: ['websocket', 'polling'] });
    socket.on('connect', () => {
      setComStatus('connected');
      // 초기 진입 시 이미 감지된 문서 목록 fetch
      fetch('/api/documents').then(r => r.json()).then(docs => setDocuments(docs)).catch(() => {});
      // 환경 상태 fetch
      fetch('/api/status').then(r => r.json()).then(s => setEnvStatus(s)).catch(() => {});
    });
    socket.on('disconnect', () => setComStatus('disconnected'));
    socket.on('documents_updated', (docs) => setDocuments(docs));
    socket.on('parse_complete', ({ doc_id }) => {
      const state = useStore.getState();
      if (state.selectedDocId === doc_id) { state.fetchParsed(doc_id); }
    });
    return () => { socket.disconnect(); socket = null; };
  }, [setDocuments, setComStatus, setEnvStatus]);
}
