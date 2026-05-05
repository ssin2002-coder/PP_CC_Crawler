import useStore from '../../stores/store.js';

// =============================================
// 규칙 카드 컴포넌트
// 각 규칙의 정보 표시 및 상호작용 처리
// =============================================

export default function RuleCard({ rule, onToggle }) {
  const activeRule = useStore((s) => s.activeRule);
  const setActiveRule = useStore((s) => s.setActiveRule);
  const setEditingRule = useStore((s) => s.setEditingRule);

  const isActive = activeRule === rule.id;

  // 왼쪽 보더 색상 결정 (심각도 기반)
  function getBorderClass() {
    if (rule.issueCount > 0) {
      if (rule.severity === 'error') return 'has-error';
      if (rule.severity === 'warning') return 'has-warning';
    }
    if (rule.type === 'auto') return 'auto-detected';
    return '';
  }

  // 카드 단일 클릭: 규칙 활성화
  function handleClick() {
    setActiveRule(rule.id);
  }

  // 카드 더블클릭: 규칙 편집 모달 열기
  function handleDoubleClick(e) {
    e.preventDefault();
    setEditingRule(rule);
  }

  // 토글 클릭: 규칙 활성화/비활성화 (이벤트 버블링 방지)
  function handleToggle(e) {
    e.stopPropagation();
    onToggle(rule.id);
  }

  // 메타 태그 렌더링 함수
  function renderMetaTags() {
    return (
      <div className="rule-card-meta">
        {/* 자동/수동 태그 */}
        <span className={`meta-tag ${rule.type === 'auto' ? 'tag-auto' : 'tag-manual'}`}>
          {rule.type === 'auto' ? '자동' : '수동'}
        </span>

        {/* 이슈 수 태그 */}
        {rule.issueCount > 0 ? (
          <span className={`meta-tag ${rule.severity === 'error' ? 'tag-error' : 'tag-warn'}`}>
            {rule.issueCount}건 적발
          </span>
        ) : (
          <span className="meta-tag tag-count">0건</span>
        )}
      </div>
    );
  }

  return (
    <div
      className={`rule-card ${getBorderClass()} ${isActive ? 'active' : ''}`}
      onClick={handleClick}
      onDoubleClick={handleDoubleClick}
      title="클릭: 규칙 선택 | 더블클릭: 편집"
    >
      <div className="rule-card-top">
        <span className="rule-card-name">{rule.name}</span>
        {/* 활성화 토글 */}
        <div
          className={`rule-card-toggle ${rule.enabled ? 'on' : ''}`}
          onClick={handleToggle}
          title={rule.enabled ? '비활성화' : '활성화'}
        />
      </div>

      {/* 설명 */}
      {rule.description && (
        <div className="rule-card-desc">{rule.description}</div>
      )}

      {/* 메타 태그 */}
      {renderMetaTags()}
    </div>
  );
}
