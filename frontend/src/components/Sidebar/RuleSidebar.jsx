import { useState } from 'react';
import useStore from '../../stores/store.js';
import { useRules } from '../../hooks/useRules.js';
import RuleTabs from './RuleTabs.jsx';
import RuleCard from './RuleCard.jsx';

// =============================================
// 규칙 사이드바 컴포넌트
// 검증 규칙 목록 표시 및 필터링
// =============================================

export default function RuleSidebar() {
  const [activeTab, setActiveTab] = useState('all');
  const setEditingRule = useStore((s) => s.setEditingRule);
  const setRuleEditorOpen = useStore((s) => s.setRuleEditorOpen);
  const { rules, toggleRule } = useRules();

  // 탭에 따른 규칙 필터링
  const filteredRules = rules.filter((rule) => {
    if (activeTab === 'all') return true;
    if (activeTab === 'auto') return rule.type === 'auto';
    if (activeTab === 'manual') return rule.type === 'manual';
    return true;
  });

  // 오류가 있는 규칙 먼저, 그 다음 경고, 그 다음 정상 순으로 정렬
  const sortedRules = [...filteredRules].sort((a, b) => {
    const order = { error: 0, warning: 1, warn: 1, info: 2 };
    const aOrder = a.issueCount > 0 ? (order[a.severity] ?? 3) : 4;
    const bOrder = b.issueCount > 0 ? (order[b.severity] ?? 3) : 4;
    return aOrder - bOrder;
  });

  // 새 규칙 추가 버튼 클릭
  function handleAddRule() {
    setEditingRule(null);    // 편집 대상 없음 = 신규 생성
    setRuleEditorOpen(true);
  }

  return (
    <div className="sidebar">
      {/* 사이드바 헤더 */}
      <div className="sidebar-header">
        <span className="sidebar-title">검증 규칙</span>
        <button
          className="btn-add-rule"
          onClick={handleAddRule}
          title="새 규칙 추가"
        >
          +
        </button>
      </div>

      {/* 탭 필터 */}
      <RuleTabs
        rules={rules}
        activeTab={activeTab}
        onTabChange={setActiveTab}
      />

      {/* 규칙 목록 */}
      <div className="rules-list">
        {sortedRules.length === 0 ? (
          <div style={{ padding: '20px', textAlign: 'center', color: '#8b949e', fontSize: '12px' }}>
            {activeTab === 'all' ? '규칙이 없습니다.' : `${activeTab === 'auto' ? '자동' : '수동'} 규칙이 없습니다.`}
          </div>
        ) : (
          sortedRules.map((rule) => (
            <RuleCard
              key={rule.id}
              rule={rule}
              onToggle={toggleRule}
            />
          ))
        )}
      </div>
    </div>
  );
}
