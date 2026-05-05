import useStore from '../../stores/store.js';

export default function IssueCard({ issue }) {
  const selectedIssue = useStore((s) => s.selectedIssue);
  const setSelectedIssue = useStore((s) => s.setSelectedIssue);

  const isActive = selectedIssue === issue.id;
  const isError = issue.severity === 'error';

  // snake_case 필드 대응 (백엔드 응답 그대로)
  const cellRef = issue.cell_ref || issue.cellRef || '';
  const ruleName = issue.rule_name || issue.ruleName || '';
  const currentValue = issue.current_value ?? issue.currentValue;
  const expectedValue = issue.expected_value ?? issue.expectedValue;

  function handleClick() {
    setSelectedIssue(issue.id);
    if (window.__navigateToCell && cellRef) {
      window.__navigateToCell(cellRef);
    }
  }

  const hasValues = isError && (currentValue != null || expectedValue != null);

  return (
    <div
      className={`ic ${isError ? 'err' : 'wrn'}${isActive ? ' active' : ''}`}
      onClick={handleClick}
    >
      <div className="ic-top">
        <span className={`ic-sev ${isError ? 'e' : 'w'}`}>
          {isError ? 'ERROR' : 'WARN'}
        </span>
        <span className="ic-cell">{cellRef}</span>
        <span className="ic-rule">{ruleName}</span>
      </div>

      <div className="ic-desc">{issue.message}</div>

      {hasValues && (
        <div className="ic-vals">
          {currentValue != null && (
            <span>
              <span className="l">현재</span>{' '}
              <span className="vr">{String(currentValue)}</span>
            </span>
          )}
          {expectedValue != null && (
            <span>
              <span className="l">기대</span>{' '}
              <span className="vg">{String(expectedValue)}</span>
            </span>
          )}
        </div>
      )}
    </div>
  );
}
