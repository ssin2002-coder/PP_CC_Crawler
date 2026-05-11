import { useStore } from '../stores/store';
const styles = {
  bar: {
    display: 'flex', justifyContent: 'space-between', alignItems: 'center',
    background: 'var(--bg-panel)', padding: '10px 20px',
    borderBottom: '1px solid var(--border)', flexShrink: 0,
  },
  left: { display: 'flex', alignItems: 'center', gap: '8px' },
  title: { fontSize: '15px', fontWeight: 600, color: '#fff', letterSpacing: '-0.015em' },
  version: { fontSize: '11px', color: 'var(--text-sub)' },
  right: { display: 'flex', alignItems: 'center', gap: '12px' },
  dot: (connected) => ({
    width: '8px', height: '8px', borderRadius: '50%',
    background: connected ? 'var(--color-green)' : '#ff453a',
  }),
  status: { fontSize: '12px', color: 'var(--text-sub)' },
  warn: { fontSize: '10px', color: 'var(--color-orange)', padding: '2px 8px',
    background: 'rgba(255, 159, 10, 0.12)', borderRadius: '4px' },
  btnPrimary: {
    background: 'var(--accent-blue)', color: '#fff', border: 'none',
    borderRadius: 'var(--radius-pill)', padding: '5px 14px',
    fontSize: '12px', cursor: 'pointer', fontWeight: 500,
  },
  btnNormal: {
    background: 'var(--bg-card)', color: 'var(--text-main)', border: '1px solid var(--border)',
    borderRadius: 'var(--radius-pill)', padding: '5px 14px',
    fontSize: '12px', cursor: 'pointer',
  },
};
export default function TopBar() {
  const documents = useStore((s) => s.documents);
  const comStatus = useStore((s) => s.comStatus);
  const envStatus = useStore((s) => s.envStatus);
  const connected = comStatus === 'connected';
  return (
    <div style={styles.bar}>
      <div style={styles.left}>
        <span style={styles.title}>Doc Intelligence</span>
        <span style={styles.version}>v0.2</span>
        {envStatus && !envStatus.tesseract_available && (
          <span style={styles.warn}>Tesseract 미설치</span>
        )}
        {envStatus && !envStatus.acrobat_available && (
          <span style={styles.warn}>Acrobat COM 미사용{envStatus.pypdf_available ? ' (pypdf 대체)' : ''}</span>
        )}
        {envStatus && !envStatus.com_available && (
          <span style={styles.warn}>COM 미사용</span>
        )}
      </div>
      <div style={styles.right}>
        <span style={styles.dot(connected)} />
        <span style={styles.status}>
          {connected ? 'COM 연결됨' : 'COM 연결 끊김'} | 문서 {documents.length}개 열림
        </span>
        <button style={styles.btnPrimary}>+ 영역 연결</button>
        <button style={styles.btnNormal}>설정</button>
      </div>
    </div>
  );
}
