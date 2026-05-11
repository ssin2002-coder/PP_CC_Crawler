const previewReasons = {
  'AcroExch.App': 'PDF — COM 미리보기 미지원',
  'Image': '이미지 파일 — 윈도우 캡처 불가',
};
export default function DocPreview({ docId, hasPreview, app }) {
  if (!hasPreview) {
    const reason = previewReasons[app] || '미리보기 없음 — 윈도우 캡처 실패';
    return (
      <div style={{
        background: '#000', borderRadius: '8px', height: '80px',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        border: '1px solid var(--border)', fontSize: '11px', color: 'var(--text-sub)',
      }}>
        {reason}
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
