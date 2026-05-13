import TopBar from './components/TopBar';
import FileList from './components/FileList';
import FingerInfo from './components/FingerInfo';
import MarkdownView, { buildMarkdown } from './components/MarkdownView';
import { useStore } from './stores/store';

function DetailPreview({ docId, documents }) {
  const doc = documents.find(d => d.id === docId);
  if (!doc || !doc.has_preview) return null;
  return (
    <div style={{
      borderBottom: '1px solid var(--border)', padding: '8px',
      background: 'var(--bg-panel)', flexShrink: 0,
      height: '200px', overflow: 'hidden',
    }}>
      <div style={{ fontSize: '11px', color: 'var(--text-sub)', marginBottom: '4px', fontWeight: 500 }}>
        원본 미리보기
      </div>
      <img src={`/api/documents/${docId}/preview`} alt="원본 미리보기"
        style={{ width: '100%', height: 'calc(100% - 20px)', objectFit: 'contain',
          borderRadius: '6px', border: '1px solid var(--border)', display: 'block' }}
        onError={(e) => { e.target.style.display = 'none'; }} />
    </div>
  );
}

function stripExt(name) {
  if (!name) return '문서';
  const idx = name.lastIndexOf('.');
  return idx > 0 ? name.slice(0, idx) : name;
}

function downloadBlob(filename, text, mime) {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function ExportToolbar({ docName, parsedData }) {
  const canExport = !!parsedData;
  const handleExport = (ext, mime) => {
    if (!canExport) return;
    const md = buildMarkdown(parsedData, docName || '문서');
    const base = stripExt(docName || '문서');
    downloadBlob(`${base}.${ext}`, md, mime);
  };
  const btn = {
    padding: '4px 10px', fontSize: '11px',
    background: 'var(--bg-card)', color: 'var(--text-main)',
    border: '1px solid var(--border)', borderRadius: 'var(--radius-pill)',
    cursor: canExport ? 'pointer' : 'not-allowed',
    opacity: canExport ? 1 : 0.5, fontWeight: 500,
  };
  return (
    <div style={{
      display: 'flex', alignItems: 'center', gap: '8px',
      padding: '6px 12px', borderBottom: '1px solid var(--border)',
      background: 'var(--bg-panel)', flexShrink: 0,
    }}>
      <span style={{
        fontSize: '12px', color: 'var(--text-main)', fontWeight: 500,
        overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
        flex: 1, minWidth: 0,
      }}>{docName || '문서'}</span>
      <button onClick={() => handleExport('md', 'text/markdown;charset=utf-8')} disabled={!canExport} style={btn}>
        📄 .md 내보내기
      </button>
      <button onClick={() => handleExport('txt', 'text/plain;charset=utf-8')} disabled={!canExport} style={btn}>
        📝 .txt 내보내기
      </button>
    </div>
  );
}

const pulseKeyframesApp = `@keyframes app-pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.5; } }`;

function PendingPane({ doc }) {
  const parseDocument = useStore((s) => s.parseDocument);
  const ps = doc.parsed_state;
  let content;
  if (ps === 'parsing') {
    content = (
      <div style={{ animation: 'app-pulse 1.2s ease-in-out infinite', fontSize: '14px', color: 'var(--text-sub)' }}>
        파싱 중...
      </div>
    );
  } else if (ps === 'error') {
    content = (
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '12px' }}>
        <div style={{ fontSize: '14px', color: '#ff453a', textAlign: 'center', maxWidth: '480px' }}>
          ⚠ 파싱 실패 — {doc.error || '알 수 없는 오류'}
        </div>
        <button onClick={() => { parseDocument(doc.id).catch(() => {}); }} style={{
          padding: '6px 16px', fontSize: '12px',
          background: 'var(--accent-blue)', color: '#fff',
          border: 'none', borderRadius: 'var(--radius-pill)', cursor: 'pointer', fontWeight: 500,
        }}>다시 시도</button>
      </div>
    );
  } else {
    content = (
      <div style={{ fontSize: '14px', color: 'var(--text-sub)' }}>
        클릭하면 파싱을 시작합니다
      </div>
    );
  }
  return (
    <div style={{
      flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center',
      flexDirection: 'column', padding: '24px',
    }}>
      <style>{pulseKeyframesApp}</style>
      {content}
    </div>
  );
}

export default function App() {
  const selectedDocId = useStore((s) => s.selectedDocId);
  const documents = useStore((s) => s.documents);
  const parsedData = useStore((s) => (selectedDocId ? s.parsedData[selectedDocId] : null));
  const selectedDoc = documents.find((d) => d.id === selectedDocId);
  const docName = selectedDoc?.name || '';
  const isParsed = selectedDoc?.parsed_state === 'parsed';
  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <TopBar />
      <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
        <div style={{ width: '33.3%', borderRight: '1px solid var(--border)', overflow: 'auto' }}>
          <FileList />
        </div>
        <div style={{ width: '66.7%', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
          {selectedDocId && selectedDoc ? (
            isParsed ? (
              <>
                <FingerInfo />
                <DetailPreview docId={selectedDocId} documents={documents} />
                <ExportToolbar docName={docName} parsedData={parsedData} />
                <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', minHeight: 0 }}>
                  <MarkdownView parsedData={parsedData} docName={docName} />
                </div>
              </>
            ) : (
              <PendingPane doc={selectedDoc} />
            )
          ) : (
            <div style={{
              flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center',
              color: 'var(--text-sub)', fontSize: '14px'
            }}>
              좌측에서 문서를 선택하세요
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
