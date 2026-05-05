import { useState, useRef, useCallback } from 'react';
import useStore from '../../stores/store.js';
import IssueTable from './IssueTable.jsx';

// =============================================
// 하단 패널 컴포넌트
// 이슈 목록 / 규칙 상세 / 검증 로그 탭 + 드래그 리사이즈
// =============================================

const MIN_HEIGHT = 80;
const MAX_HEIGHT = 500;
const DEFAULT_HEIGHT = 180;

const TABS = ['이슈 목록', '규칙 상세', '검증 로그'];

export default function BottomPanel() {
  const [activeTab, setActiveTab] = useState('이슈 목록');
  const [panelHeight, setPanelHeight] = useState(DEFAULT_HEIGHT);

  const issues = useStore((s) => s.issues);
  const activeRule = useStore((s) => s.activeRule);
  const rules = useStore((s) => s.rules);

  const dragRef = useRef({ isDragging: false, startY: 0, startHeight: 0 });

  // 이슈 수 계산 (활성 규칙 필터 적용)
  const displayedIssueCount = activeRule
    ? issues.filter((i) => i.ruleId === activeRule).length
    : issues.length;

  // 드래그로 패널 높이 조정
  const handleDragStart = useCallback((e) => {
    dragRef.current = {
      isDragging: true,
      startY: e.clientY,
      startHeight: panelHeight,
    };

    function onMouseMove(e) {
      if (!dragRef.current.isDragging) return;
      const delta = dragRef.current.startY - e.clientY; // 위로 드래그 = 높이 증가
      const newHeight = Math.min(
        MAX_HEIGHT,
        Math.max(MIN_HEIGHT, dragRef.current.startHeight + delta)
      );
      setPanelHeight(newHeight);
    }

    function onMouseUp() {
      dragRef.current.isDragging = false;
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    }

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }, [panelHeight]);

  // 활성 규칙 상세 정보
  const activeRuleObj = activeRule ? rules.find((r) => r.id === activeRule) : null;

  return (
    <div
      className="bottom-panel"
      style={{ height: panelHeight }}
    >
      {/* 드래그 핸들 */}
      <div
        className="bottom-drag-handle"
        onMouseDown={handleDragStart}
        title="드래그하여 패널 높이 조절"
      />

      {/* 탭 헤더 */}
      <div className="bottom-tabs">
        {TABS.map((tab) => (
          <div
            key={tab}
            className={`bottom-tab${activeTab === tab ? ' active' : ''}`}
            onClick={() => setActiveTab(tab)}
          >
            {tab}
            {/* 이슈 수 배지 (이슈 목록 탭만) */}
            {tab === '이슈 목록' && displayedIssueCount > 0 && (
              <span className="count">{displayedIssueCount}</span>
            )}
          </div>
        ))}
      </div>

      {/* 탭 콘텐츠 */}
      <div className="bottom-content">
        {activeTab === '이슈 목록' && <IssueTable />}

        {activeTab === '규칙 상세' && (
          <div style={{ padding: '12px 14px' }}>
            {activeRuleObj ? (
              <>
                <div style={{ fontSize: '13px', fontWeight: 'bold', color: '#c9d1d9', marginBottom: '8px' }}>
                  {activeRuleObj.name}
                </div>
                <div style={{ fontSize: '11px', color: '#8b949e', marginBottom: '6px' }}>
                  {activeRuleObj.description}
                </div>
                <div style={{ display: 'flex', gap: '8px', fontSize: '10px' }}>
                  <span className={`meta-tag ${activeRuleObj.type === 'auto' ? 'tag-auto' : 'tag-manual'}`}>
                    {activeRuleObj.type === 'auto' ? '자동' : '수동'}
                  </span>
                  <span className={`meta-tag ${activeRuleObj.severity === 'error' ? 'tag-error' : 'tag-warn'}`}>
                    {activeRuleObj.severity === 'error' ? '오류' : '경고'}
                  </span>
                </div>
              </>
            ) : (
              <div style={{ color: '#8b949e', fontSize: '12px' }}>
                좌측에서 규칙을 선택하면 상세 정보가 표시됩니다.
              </div>
            )}
          </div>
        )}

        {activeTab === '검증 로그' && (
          <div style={{ padding: '12px 14px', fontSize: '11px', color: '#8b949e' }}>
            검증 로그가 여기에 표시됩니다.
          </div>
        )}
      </div>
    </div>
  );
}
