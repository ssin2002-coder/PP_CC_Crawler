import { useSocket } from '../hooks/useSocket';
import { useStore } from '../stores/store';
import FileCard from './FileCard';
export default function FileList() {
  useSocket();
  const documents = useStore((s) => s.documents);
  const selectedDocId = useStore((s) => s.selectedDocId);
  const selectDocument = useStore((s) => s.selectDocument);
  const fetchParsed = useStore((s) => s.fetchParsed);
  const handleSelect = (docId) => { selectDocument(docId); fetchParsed(docId); };
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
