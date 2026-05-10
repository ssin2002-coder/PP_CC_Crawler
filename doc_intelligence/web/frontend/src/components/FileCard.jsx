import DocPreview from './DocPreview';
const iconMap = {
  'Excel.Application': '📊', 'Word.Application': '📝', 'PowerPoint.Application': '📑',
};
const statusConfig = {
  matched: { label: '매칭됨', bg: 'var(--color-green)', color: '#000' },
  candidate: { label: '후보', bg: 'var(--color-orange)', color: '#000' },
  new: { label: '새 문서', bg: 'var(--accent-blue-light)', color: '#fff' },
};
export default function FileCard({ doc, selected, onClick }) {
  const icon = iconMap[doc.app] || '📄';
  const badge = statusConfig[doc.status] || statusConfig.new;
  return (
    <div onClick={onClick} style={{
      background: selected ? 'rgba(0, 113, 227, 0.15)' : 'var(--bg-card)',
      border: `1px solid ${selected ? 'var(--accent-blue)' : 'var(--border)'}`,
      borderRadius: 'var(--radius-card)', padding: '10px',
      marginBottom: '8px', cursor: 'pointer', transition: 'border-color 0.2s',
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' }}>
        <span style={{ fontSize: '16px' }}>{icon}</span>
        <span style={{ fontSize: '12px', color: 'var(--text-main)', fontWeight: 500,
          overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          {doc.name}
        </span>
      </div>
      <DocPreview docId={doc.id} hasPreview={doc.has_preview} />
      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '6px' }}>
        <span style={{ fontSize: '10px', background: badge.bg, color: badge.color,
          padding: '1px 8px', borderRadius: 'var(--radius-pill)', fontWeight: 500 }}>
          {badge.label}
        </span>
        {doc.template_name && (
          <span style={{ fontSize: '10px', color: 'var(--text-sub)' }}>
            {doc.template_name} ({doc.score}%)
          </span>
        )}
      </div>
    </div>
  );
}
