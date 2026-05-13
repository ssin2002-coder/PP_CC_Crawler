import { useState } from 'react';
import { useStore } from '../stores/store';
export default function FingerInfo() {
  const selectedDocId = useStore((s) => s.selectedDocId);
  const documents = useStore((s) => s.documents);
  const refetchDocuments = useStore((s) => s.refetchDocuments);
  const [learnName, setLearnName] = useState('');
  const [showModal, setShowModal] = useState(false);
  const [busy, setBusy] = useState(false);
  const [feedback, setFeedback] = useState(null);
  const doc = documents.find((d) => d.id === selectedDocId);
  if (!doc) return null;

  const showFeedback = (kind, message) => {
    setFeedback({ kind, message });
    setTimeout(() => setFeedback(null), 3000);
  };

  const handleLearn = async () => {
    const name = learnName.trim();
    if (!name || busy) return;
    setBusy(true);
    try {
      const res = await fetch('/api/templates/learn', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ doc_id: doc.id, template_name: name }),
      });
      if (!res.ok) throw new Error(`learn failed: ${res.status}`);
      await refetchDocuments();
      setShowModal(false);
      setLearnName('');
      showFeedback('ok', `'${name}' 템플릿으로 학습되었습니다`);
    } catch (e) {
      showFeedback('err', `학습 실패: ${e.message}`);
    } finally {
      setBusy(false);
    }
  };
  const handleConfirm = async () => {
    if (busy) return;
    setBusy(true);
    try {
      const res = await fetch('/api/templates/confirm', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ doc_id: doc.id, template_id: doc.template_id }),
      });
      if (!res.ok) throw new Error(`confirm failed: ${res.status}`);
      await refetchDocuments();
      showFeedback('ok', '템플릿이 확정되었습니다');
    } catch (e) {
      showFeedback('err', `확정 실패: ${e.message}`);
    } finally {
      setBusy(false);
    }
  };
  return (
    <div style={{ background: 'var(--bg-panel)', borderBottom: '1px solid var(--border)', padding: '12px 16px' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
        {doc.status === 'matched' && (
          <span style={{ fontSize: '13px', color: 'var(--color-green)' }}>✓ {doc.template_name} ({doc.score}%)</span>
        )}
        {doc.status === 'candidate' && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '13px', color: 'var(--color-orange)' }}>? 이 양식인가요: {doc.template_name} ({doc.score}%)</span>
            <button onClick={handleConfirm} disabled={busy} style={{
              background: 'var(--color-green)', color: '#000', border: 'none',
              borderRadius: 'var(--radius-pill)', padding: '3px 12px', fontSize: '11px',
              cursor: busy ? 'wait' : 'pointer', fontWeight: 500, opacity: busy ? 0.6 : 1,
            }}>{busy ? '...' : '예'}</button>
            <button onClick={() => setShowModal(true)} disabled={busy} style={{
              background: 'var(--bg-card)', color: 'var(--text-main)', border: '1px solid var(--border)',
              borderRadius: 'var(--radius-pill)', padding: '3px 12px', fontSize: '11px',
              cursor: busy ? 'wait' : 'pointer', opacity: busy ? 0.6 : 1,
            }}>아니오</button>
          </div>
        )}
        {doc.status === 'new' && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '13px', color: 'var(--accent-blue-light)' }}>새 문서 — 템플릿 학습 필요</span>
            <button onClick={() => setShowModal(true)} disabled={busy} style={{
              background: 'var(--accent-blue)', color: '#fff', border: 'none',
              borderRadius: 'var(--radius-pill)', padding: '3px 12px', fontSize: '11px',
              cursor: busy ? 'wait' : 'pointer', fontWeight: 500, opacity: busy ? 0.6 : 1,
            }}>학습</button>
          </div>
        )}
        {feedback && (
          <span style={{
            fontSize: '11px', marginLeft: 'auto',
            color: feedback.kind === 'ok' ? 'var(--color-green)' : '#ff453a',
          }}>{feedback.kind === 'ok' ? '✓' : '⚠'} {feedback.message}</span>
        )}
      </div>
      {doc.labels && doc.labels.length > 0 && (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px' }}>
          {doc.labels.slice(0, 20).map((label, i) => (
            <span key={i} style={{
              fontSize: '10px', background: 'rgba(255,255,255,0.08)', color: 'var(--text-sub)',
              padding: '2px 8px', borderRadius: 'var(--radius-pill)', border: '1px solid var(--border)',
            }}>{label}</span>
          ))}
          {doc.labels.length > 20 && (
            <span style={{ fontSize: '10px', color: 'var(--text-sub)' }}>+{doc.labels.length - 20}</span>
          )}
        </div>
      )}
      {showModal && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.7)',
          display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000,
        }}>
          <div style={{
            background: 'var(--bg-panel)', borderRadius: 'var(--radius-card)',
            padding: '24px', width: '360px', border: '1px solid var(--border)',
          }}>
            <h3 style={{ fontSize: '15px', marginBottom: '12px', fontWeight: 600 }}>템플릿 학습</h3>
            <input value={learnName} onChange={(e) => setLearnName(e.target.value)}
              placeholder="양식 이름 입력 (예: 정비비용정산서)"
              style={{
                width: '100%', padding: '8px 12px', borderRadius: '8px',
                border: '1px solid var(--border)', background: '#000',
                color: 'var(--text-main)', fontSize: '13px', outline: 'none',
              }}
              onKeyDown={(e) => e.key === 'Enter' && handleLearn()} autoFocus />
            <div style={{ display: 'flex', gap: '8px', marginTop: '16px', justifyContent: 'flex-end' }}>
              <button onClick={() => setShowModal(false)} disabled={busy} style={{
                background: 'var(--bg-card)', color: 'var(--text-main)', border: '1px solid var(--border)',
                borderRadius: 'var(--radius-pill)', padding: '6px 16px', fontSize: '12px',
                cursor: busy ? 'wait' : 'pointer', opacity: busy ? 0.6 : 1,
              }}>취소</button>
              <button onClick={handleLearn} disabled={busy || !learnName.trim()} style={{
                background: 'var(--accent-blue)', color: '#fff', border: 'none',
                borderRadius: 'var(--radius-pill)', padding: '6px 16px', fontSize: '12px',
                cursor: (busy || !learnName.trim()) ? 'not-allowed' : 'pointer',
                fontWeight: 500, opacity: (busy || !learnName.trim()) ? 0.6 : 1,
              }}>{busy ? '학습 중...' : '학습'}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
