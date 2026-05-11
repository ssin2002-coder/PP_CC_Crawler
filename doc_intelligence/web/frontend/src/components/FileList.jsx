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
  const setDocuments = useStore((s) => s.setDocuments);
  const [uploadPath, setUploadPath] = useState('');
  const [uploadType, setUploadType] = useState(null);
  const [uploading, setUploading] = useState(false);

  const handleSelect = (docId) => { selectDocument(docId); fetchParsed(docId); };

  const handleUpload = async () => {
    if (!uploadPath.trim() || !uploadType) return;
    setUploading(true);
    try {
      const endpoint = uploadType === 'pdf' ? '/api/documents/upload-pdf' : '/api/documents/upload-image';
      const res = await fetch(endpoint, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_path: uploadPath.trim() }),
      });
      if (res.ok) {
        const result = await res.json();
        const docsRes = await fetch('/api/documents');
        if (docsRes.ok) setDocuments(await docsRes.json());
        if (result.doc_id) { selectDocument(result.doc_id); fetchParsed(result.doc_id); }
      }
    } catch (e) { console.error('Upload failed:', e); }
    setUploading(false);
    setUploadPath('');
    setUploadType(null);
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

      {/* 업로드 영역 */}
      {uploadType ? (
        <div style={{ padding: '8px', borderBottom: '1px solid var(--border)', background: 'var(--bg-panel)' }}>
          <div style={{ fontSize: '11px', color: 'var(--text-sub)', marginBottom: '4px' }}>
            {uploadType === 'pdf' ? 'PDF' : '이미지'} 파일 경로
          </div>
          <div style={{ display: 'flex', gap: '4px' }}>
            <input value={uploadPath} onChange={(e) => setUploadPath(e.target.value)}
              placeholder="C:\path\to\file.pdf"
              onKeyDown={(e) => e.key === 'Enter' && handleUpload()}
              autoFocus
              style={{
                flex: 1, padding: '6px 8px', fontSize: '11px', borderRadius: '6px',
                border: '1px solid var(--border)', background: '#000',
                color: 'var(--text-main)', outline: 'none',
              }} />
            <button onClick={handleUpload} disabled={uploading} style={{
              padding: '6px 12px', fontSize: '11px', cursor: 'pointer',
              background: 'var(--accent-blue)', color: '#fff',
              border: 'none', borderRadius: '6px', fontWeight: 500,
            }}>{uploading ? '...' : '추가'}</button>
            <button onClick={() => { setUploadType(null); setUploadPath(''); }} style={{
              padding: '6px 8px', fontSize: '11px', cursor: 'pointer',
              background: 'var(--bg-card)', color: 'var(--text-sub)',
              border: '1px solid var(--border)', borderRadius: '6px',
            }}>취소</button>
          </div>
        </div>
      ) : (
        <div style={{ display: 'flex', gap: '4px', padding: '8px 8px 0' }}>
          <button onClick={() => setUploadType('pdf')} style={{
            flex: 1, padding: '6px 0', fontSize: '11px', cursor: 'pointer',
            background: 'var(--bg-card)', color: 'var(--text-main)',
            border: '1px solid var(--border)', borderRadius: 'var(--radius-pill)',
          }}>+ PDF</button>
          <button onClick={() => setUploadType('image')} style={{
            flex: 1, padding: '6px 0', fontSize: '11px', cursor: 'pointer',
            background: 'var(--bg-card)', color: 'var(--text-main)',
            border: '1px solid var(--border)', borderRadius: 'var(--radius-pill)',
          }}>+ 이미지</button>
        </div>
      )}

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
