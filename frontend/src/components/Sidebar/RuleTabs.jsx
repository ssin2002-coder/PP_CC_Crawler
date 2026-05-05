// =============================================
// 규칙 탭 필터 컴포넌트
// 전체 / 자동 / 수동 탭으로 규칙 필터링
// =============================================

const TABS = [
  { id: 'all', label: '전체' },
  { id: 'auto', label: '자동' },
  { id: 'manual', label: '수동' },
];

export default function RuleTabs({ rules, activeTab, onTabChange }) {
  // 각 탭별 규칙 수 계산
  const counts = {
    all: rules.length,
    auto: rules.filter((r) => r.type === 'auto').length,
    manual: rules.filter((r) => r.type === 'manual').length,
  };

  return (
    <div className="sidebar-tabs">
      {TABS.map((tab) => (
        <div
          key={tab.id}
          className={`sidebar-tab${activeTab === tab.id ? ' active' : ''}`}
          onClick={() => onTabChange(tab.id)}
        >
          {tab.label} ({counts[tab.id]})
        </div>
      ))}
    </div>
  );
}
