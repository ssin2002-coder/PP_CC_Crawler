import DocPreview from './DocPreview';
const iconMap = {
  'Excel.Application': '📊', 'Word.Application': '📝', 'PowerPoint.Application': '📑',
  'AcroExch.App': '📕', 'Image': '🖼️',
};
const statusConfig = {
  matched: { label: '매칭됨', bg: 'var(--color-green)', color: '#000' },
  candidate: { label: '후보', bg: 'var(--color-orange)', color: '#000' },
  new: { label: '새 문서', bg: 'var(--accent-blue-light)', color: '#fff' },
};

const pulseKeyframes = `@keyframes fc-pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.55; } }`;

function StateBadge({ doc }) {
  const ps = doc.parsed_state;
  if (ps === 'discovered') {
    return (
      <span style={{ fontSize: '10px', background: 'var(--border)', color: 'var(--text-main)',
        padding: '1px 8px', borderRadius: 'var(--radius-pill)', fontWeight: 500 }}>
        대기
      </span>
    );
  }
  if (ps === 'parsing') {
    return (
      <span style={{ fontSize: '10px', background: 'var(--color-orange)', color: '#000',
        padding: '1px 8px', borderRadius: 'var(--radius-pill)', fontWeight: 500,
        animation: 'fc-pulse 1.2s ease-in-out infinite' }}>
        파싱 중...
      </span>
    );
  }
  if (ps === 'error') {
    return (
      <span title={doc.error || ''} style={{ fontSize: '10px', background: '#ff453a', color: '#fff',
        padding: '1px 8px', borderRadius: 'var(--radius-pill)', fontWeight: 500, cursor: 'help' }}>
        오류
      </span>
    );
  }
  const badge = statusConfig[doc.status] || statusConfig.new;
  return (
    <>
      <span style={{ fontSize: '10px', background: badge.bg, color: badge.color,
        padding: '1px 8px', borderRadius: 'var(--radius-pill)', fontWeight: 500 }}>
        {badge.label}
      </span>
      {doc.template_name && (
        <span style={{ fontSize: '10px', color: 'var(--text-sub)' }}>
          {doc.template_name} ({doc.score}%)
        </span>
      )}
    </>
  );
}

export default function FileCard({ doc, selected, onClick }) {
  const icon = iconMap[doc.app] || '📄';
  return (
    <div onClick={onClick} style={{
      background: selected ? 'rgba(0, 113, 227, 0.15)' : 'var(--bg-card)',
      border: `1px solid ${selected ? 'var(--accent-blue)' : 'var(--border)'}`,
      borderRadius: 'var(--radius-card)', padding: '10px',
      marginBottom: '8px', cursor: 'pointer', transition: 'border-color 0.2s',
    }}>
      <style>{pulseKeyframes}</style>
      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' }}>
        <span style={{ fontSize: '16px' }}>{icon}</span>
        <span style={{ fontSize: '12px', color: 'var(--text-main)', fontWeight: 500,
          overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          {doc.name}
        </span>
      </div>
      <DocPreview docId={doc.id} hasPreview={doc.has_preview} app={doc.app} />
      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '6px' }}>
        <StateBadge doc={doc} />
      </div>
    </div>
  );
}
