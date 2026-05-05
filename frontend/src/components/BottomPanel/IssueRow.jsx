import useStore from '../../stores/store.js';

// =============================================
// 이슈 행 컴포넌트
// 이슈 목록 테이블의 각 행을 렌더링
// =============================================

export default function IssueRow({ issue }) {
  const selectedIssue = useStore((s) => s.selectedIssue);
  const setSelectedIssue = useStore((s) => s.setSelectedIssue);

  const isActive = selectedIssue === issue.id;
  const isError = issue.severity === 'error';

  // 행 클릭: 이슈 선택 + 셀 이동 + 규칙 활성화 (스토어에서 3-way sync 처리)
  function handleClick() {
    setSelectedIssue(issue.id);

    // Excel 파일로 네비게이션
    if (window.__navigateToCell) {
      window.__navigateToCell(issue.cellRef);
    }
  }

  return (
    <tr
      className={isActive ? 'active-row' : ''}
      onClick={handleClick}
      style={{ cursor: 'pointer' }}
    >
      {/* 심각도 점 */}
      <td style={{ textAlign: 'center' }}>
        <span className={`severity-dot ${isError ? 'dot-error' : 'dot-warn'}`} />
      </td>

      {/* 셀 위치 배지 */}
      <td>
        <span className="cell-ref-badge">{issue.cellRef}</span>
      </td>

      {/* 규칙 이름 */}
      <td style={{ color: '#c9d1d9' }}>{issue.ruleName}</td>

      {/* 이슈 설명 */}
      <td style={{ color: '#8b949e' }}>{issue.message}</td>

      {/* 현재값 */}
      <td style={{ color: isError ? '#f85149' : '#d29922' }}>
        {issue.currentValue ?? '-'}
      </td>

      {/* 기대값 */}
      <td style={{ color: '#3fb950' }}>
        {issue.expectedValue ?? '-'}
      </td>
    </tr>
  );
}
