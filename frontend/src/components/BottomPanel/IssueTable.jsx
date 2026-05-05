import useStore from '../../stores/store.js';
import IssueRow from './IssueRow.jsx';

// =============================================
// 이슈 테이블 컴포넌트
// 모든 이슈를 테이블 형태로 표시 (오류 우선 정렬)
// =============================================

export default function IssueTable() {
  const issues = useStore((s) => s.issues);
  const activeRule = useStore((s) => s.activeRule);

  // 활성 규칙 필터링 (규칙 선택 시 해당 규칙 이슈만 표시)
  const filteredIssues = activeRule
    ? issues.filter((i) => i.ruleId === activeRule)
    : issues;

  // 오류 먼저, 그 다음 경고 순으로 정렬
  const sortedIssues = [...filteredIssues].sort((a, b) => {
    const order = { error: 0, warning: 1, warn: 1, info: 2 };
    return (order[a.severity] ?? 3) - (order[b.severity] ?? 3);
  });

  if (sortedIssues.length === 0) {
    return (
      <div style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100px',
        color: '#8b949e',
        fontSize: '12px',
      }}>
        {activeRule ? '선택된 규칙에 이슈가 없습니다.' : '이슈가 없습니다. 검증을 실행하십시오.'}
      </div>
    );
  }

  return (
    <table className="issue-table">
      <thead>
        <tr>
          <th style={{ width: 30 }} />
          <th style={{ width: 80 }}>위치</th>
          <th style={{ width: 140 }}>규칙</th>
          <th>설명</th>
          <th style={{ width: 90 }}>현재값</th>
          <th style={{ width: 90 }}>기대값</th>
        </tr>
      </thead>
      <tbody>
        {sortedIssues.map((issue) => (
          <IssueRow key={issue.id} issue={issue} />
        ))}
      </tbody>
    </table>
  );
}
