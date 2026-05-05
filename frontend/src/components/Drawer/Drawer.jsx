import { useState } from 'react';
import useStore from '../../stores/store.js';
import IssueCard from './IssueCard.jsx';
import RuleChips from './RuleChips.jsx';

// =============================================
// Drawer 컴포넌트 (Sidebar + BottomPanel 대체)
// 왼쪽 330px 서랍 패널
// 탭: 이슈 / 규칙 / 로그
// 하단: 규칙 칩 바
// =============================================

const TABS = ['이슈', '규칙', '로그'];

export default function Drawer() {
  const [activeTab, setActiveTab] = useState('이슈');

  const issues = useStore((s) => s.issues);
  const activeRule = useStore((s) => s.activeRule);

  // 활성 규칙 필터링 (snake_case 대응)
  const displayedIssues = activeRule
    ? issues.filter((i) => (i.rule_id || i.ruleId) === activeRule)
    : issues;

  const issueCount = issues.length;

  return (
    <div className="drawer">
      {/* 탭 헤더 */}
      <div className="drawer-tabs">
        {TABS.map((tab) => (
          <div
            key={tab}
            className={`drawer-tab${activeTab === tab ? ' active' : ''}`}
            onClick={() => setActiveTab(tab)}
          >
            {tab}
            {/* 이슈 탭에만 빨간 배지 */}
            {tab === '이슈' && issueCount > 0 && (
              <span className="badge badge-red">{issueCount}</span>
            )}
          </div>
        ))}
      </div>

      {/* 탭 구분선 */}
      <div className="drawer-line" />

      {/* 탭 콘텐츠 */}
      <div className="drawer-content">
        {activeTab === '이슈' && (
          displayedIssues.length > 0 ? (
            displayedIssues.map((issue) => (
              <IssueCard key={issue.id} issue={issue} />
            ))
          ) : (
            <div className="drawer-empty">
              {issues.length === 0
                ? '이슈가 없습니다.'
                : '선택한 규칙에 이슈가 없습니다.'}
            </div>
          )
        )}

        {activeTab === '규칙' && (
          <div className="drawer-empty">
            규칙 목록은 하단 칩에서 확인하세요.
          </div>
        )}

        {activeTab === '로그' && (
          <div className="drawer-empty">
            검증 로그가 여기에 표시됩니다.
          </div>
        )}
      </div>

      {/* 하단 규칙 칩 바 */}
      <RuleChips />
    </div>
  );
}
