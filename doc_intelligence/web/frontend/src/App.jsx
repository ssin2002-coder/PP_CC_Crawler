import TopBar from './components/TopBar';
import FileList from './components/FileList';
import FingerInfo from './components/FingerInfo';
import DataTable from './components/DataTable';
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

export default function App() {
  const selectedDocId = useStore((s) => s.selectedDocId);
  const documents = useStore((s) => s.documents);
  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <TopBar />
      <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
        <div style={{ width: '33.3%', borderRight: '1px solid var(--border)', overflow: 'auto' }}>
          <FileList />
        </div>
        <div style={{ width: '66.7%', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
          {selectedDocId ? (
            <>
              <FingerInfo />
              <DetailPreview docId={selectedDocId} documents={documents} />
              <div style={{ flex: 1, overflow: 'auto', minHeight: 0 }}>
                <DataTable />
              </div>
            </>
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
