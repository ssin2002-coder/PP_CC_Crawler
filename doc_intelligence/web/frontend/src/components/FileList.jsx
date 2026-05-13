import { useState } from 'react';
import { useSocket } from '../hooks/useSocket';
import { useStore } from '../stores/store';
import FileCard from './FileCard';
export default function FileList() {
  useSocket();
  const documents = useStore((s) => s.documents);
  const selectedDocId = useStore((s) => s.selectedDocId);
  const selectDocument = useStore((s) => s.selectDocument);
  const fetchParsed = useStore((s) => s.fetchParsed);
  const refetchDocuments = useStore((s) => s.refetchDocuments);
  const scanDocuments = useStore((s) => s.scanDocuments);
  const parseDocument = useStore((s) => s.parseDocument);
  const [scanning, setScanning] = useState(false);
  const [status, setStatus] = useState(null);

  const showStatus = (kind, message, ttl = 3000) => {
    setStatus({ kind, message });
    if (ttl > 0) setTimeout(() => setStatus(null), ttl);
  };

  const handleSelect = async (docId) => {
    selectDocument(docId);
    const doc = useStore.getState().documents.find((d) => d.id === docId);
    if (doc?.parsed_state === 'parsed') {
      fetchParsed(docId);
      return;
    }
    try {
      await parseDocument(docId);
    } catch (e) {
      showStatus('err', `⚠ 파싱 실패: ${e.message}`);
    }
  };

  const handleScan = async () => {
    if (scanning) return;
    setScanning(true);
    try {
      const result = await scanDocuments();
      await refetchDocuments();
      const detected = result?.detected ?? 0;
      if (detected === 0) {
        showStatus('ok', '· 감지된 파일 없음');
      } else {
        showStatus('ok', `✓ ${detected}개 감지 — 카드를 클릭해 파싱하세요`);
      }
    } catch (e) {
      showStatus('err', `⚠ 감지 실패: ${e.message}`);
    } finally {
      setScanning(false);
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div style={{
        background: 'var(--bg-panel)', padding: '10px 14px',
        display: 'flex', justifyContent: 'space-between', alignItems: 'center',
        borderBottom: '1px solid var(--border)',
      }}>
        <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-main)' }}>열린 문서</span>
        <span style={{ fontSize: '12px', color: 'var(--accent-blue-light)' }}>{documents.length}개 감지</span>
      </div>

      <div style={{ padding: '8px 8px 0' }}>
        <button onClick={handleScan} disabled={scanning} style={{
          width: '100%', padding: '8px 0', fontSize: '12px',
          cursor: scanning ? 'wait' : 'pointer',
          background: scanning ? 'var(--bg-card)' : 'var(--accent-blue)',
          color: scanning ? 'var(--text-sub)' : '#fff',
          border: scanning ? '1px solid var(--border)' : 'none',
          borderRadius: 'var(--radius-pill)',
          fontWeight: 500,
          opacity: scanning ? 0.7 : 1,
        }}>
          {scanning ? '감지 중...' : '📂 문서+이미지 감지'}
        </button>
        {status && (
          <div style={{
            marginTop: '6px', fontSize: '11px', textAlign: 'center',
            color: status.kind === 'ok' ? 'var(--color-green)' : '#ff453a',
          }}>{status.message}</div>
        )}
      </div>

      <div style={{ flex: 1, overflow: 'auto', padding: '8px' }}>
        {documents.length === 0 ? (
          <div style={{ textAlign: 'center', color: 'var(--text-sub)', fontSize: '12px', marginTop: '40px' }}>
            열린 문서가 없습니다
          </div>
        ) : (
          documents.map((doc) => (
            <FileCard key={doc.id} doc={doc} selected={doc.id === selectedDocId}
              onClick={() => handleSelect(doc.id)} />
          ))
        )}
      </div>
    </div>
  );
}
