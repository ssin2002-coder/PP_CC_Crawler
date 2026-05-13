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
  const detectFiles = useStore((s) => s.detectFiles);
  const [detecting, setDetecting] = useState(false);
  const [status, setStatus] = useState(null);

  const handleSelect = (docId) => { selectDocument(docId); fetchParsed(docId); };

  const showStatus = (kind, message) => {
    setStatus({ kind, message });
    setTimeout(() => setStatus(null), 3000);
  };

  const handleDetect = async () => {
    if (detecting) return;
    setDetecting(true);
    try {
      const result = await detectFiles();
      await refetchDocuments();
      const detected = result?.detected ?? 0;
      const newlyParsed = result?.newly_parsed ?? 0;
      const alreadyCached = result?.already_cached ?? 0;
      if (newlyParsed === 0 && detected === 0) {
        showStatus('ok', '· 새 파일 없음');
      } else {
        showStatus('ok', `✓ ${detected}개 감지 (신규 ${newlyParsed}, 기존 ${alreadyCached})`);
      }
    } catch (e) {
      showStatus('err', `⚠ 감지 실패: ${e.message}`);
    } finally {
      setDetecting(false);
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
        <button onClick={handleDetect} disabled={detecting} style={{
          width: '100%', padding: '8px 0', fontSize: '12px',
          cursor: detecting ? 'wait' : 'pointer',
          background: detecting ? 'var(--bg-card)' : 'var(--accent-blue)',
          color: detecting ? 'var(--text-sub)' : '#fff',
          border: detecting ? '1px solid var(--border)' : 'none',
          borderRadius: 'var(--radius-pill)',
          fontWeight: 500,
          opacity: detecting ? 0.7 : 1,
        }}>
          {detecting ? '감지 중...' : '📂 문서+이미지 감지'}
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
