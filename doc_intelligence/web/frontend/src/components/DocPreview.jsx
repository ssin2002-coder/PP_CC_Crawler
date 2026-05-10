export default function DocPreview({ docId, hasPreview }) {
  if (!hasPreview) {
    return (
      <div style={{
        background: '#000', borderRadius: '8px', height: '80px',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        border: '1px solid var(--border)', fontSize: '11px', color: 'var(--text-sub)',
      }}>
        미리보기 없음
      </div>
    );
  }
  return (
    <img src={`/api/documents/${docId}/preview`} alt="preview"
      style={{ width: '100%', height: '80px', objectFit: 'cover',
        borderRadius: '8px', border: '1px solid var(--border)' }}
      onError={(e) => { e.target.style.display = 'none'; }} />
  );
}
