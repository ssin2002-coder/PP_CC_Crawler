import { useEffect, useRef } from 'react';

// =============================================
// 셀 툴팁 컴포넌트
// 오류/경고 셀 호버 시 상세 정보 표시
// =============================================

export default function CellTooltip({ issues, anchorRect, visible }) {
  const tooltipRef = useRef(null);

  useEffect(() => {
    if (!visible || !tooltipRef.current || !anchorRect) return;

    const tooltip = tooltipRef.current;
    const tooltipRect = tooltip.getBoundingClientRect();
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;

    // 기본 위치: 셀 아래쪽
    let top = anchorRect.bottom + 4;
    let left = anchorRect.left;

    // 오른쪽 넘침 방지
    if (left + tooltipRect.width > viewportWidth) {
      left = viewportWidth - tooltipRect.width - 8;
    }

    // 아래쪽 넘침 방지 - 위쪽에 표시
    if (top + tooltipRect.height > viewportHeight) {
      top = anchorRect.top - tooltipRect.height - 4;
    }

    tooltip.style.top = `${top}px`;
    tooltip.style.left = `${left}px`;
  }, [visible, anchorRect]);

  if (!visible || !issues || issues.length === 0) return null;

  // 첫 번째 이슈를 메인으로 표시 (여러 이슈일 경우 count 표시)
  const primaryIssue = issues[0];
  const isError = primaryIssue.severity === 'error';
  const borderColor = isError ? '#f85149' : '#d29922';

  return (
    <div
      ref={tooltipRef}
      className="cell-tooltip"
      style={{
        display: 'block',
        borderColor,
        position: 'fixed',
      }}
    >
      {/* 규칙 이름 */}
      <div className="tt-rule">{primaryIssue.ruleName}</div>

      {/* 이슈 메시지 */}
      <div className="tt-msg">{primaryIssue.message}</div>

      {/* 현재값 / 기대값 */}
      <div className="tt-values">
        <span>
          현재: <span className="tt-current">{primaryIssue.currentValue ?? '-'}</span>
        </span>
        <span>
          기대: <span className="tt-expected">{primaryIssue.expectedValue ?? '-'}</span>
        </span>
      </div>

      {/* 이슈가 여러 개일 경우 */}
      {issues.length > 1 && (
        <div style={{ marginTop: '4px', fontSize: '10px', color: '#8b949e' }}>
          외 {issues.length - 1}개 이슈
        </div>
      )}
    </div>
  );
}
