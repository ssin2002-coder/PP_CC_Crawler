import { useEffect } from 'react';
import { io } from 'socket.io-client';
import { useStore } from '../stores/store';
let socket = null;
export function useSocket() {
  const setDocuments = useStore((s) => s.setDocuments);
  const setComStatus = useStore((s) => s.setComStatus);
  useEffect(() => {
    socket = io({ transports: ['websocket', 'polling'] });
    socket.on('connect', () => setComStatus('connected'));
    socket.on('disconnect', () => setComStatus('disconnected'));
    socket.on('documents_updated', (docs) => setDocuments(docs));
    socket.on('parse_complete', ({ doc_id }) => {
      const state = useStore.getState();
      if (state.selectedDocId === doc_id) { state.fetchParsed(doc_id); }
    });
    return () => { socket.disconnect(); socket = null; };
  }, [setDocuments, setComStatus]);
}
