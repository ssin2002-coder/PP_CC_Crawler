import { useState, useRef } from 'react';
import CellTooltip from './CellTooltip.jsx';
import useStore from '../../stores/store.js';
import { toCellRef } from '../../utils/cellRef.js';

// =============================================
// GridCell 컴포넌트
// Excel 셀 하나의 렌더링 및 상호작용 처리
// 목업 v2 셀 클래스 적용: ce/cw/cd/cs/cl/ch/cn/rs/rt
// =============================================

export default function GridCell({ row, col, cell, isSelected, isError, isWarning, isRelated, issues }) {
  const [tooltipVisible, setTooltipVisible] = useState(false);
  const [anchorRect, setAnchorRect] = useState(null);
  const cellRef = useRef(null);

  const setSelectedCell = useStore((s) => s.setSelectedCell);

  const value = cell ? cell.value : '';
  const isBold = cell && cell.bold;
  const isHeader = cell && cell.isHeader;
  const isNumeric = typeof value === 'number' || (cell && cell.isNumeric);
  const isDuplicate = issues && issues.some((i) => i.type === 'duplicate' || i.ruleName?.includes('중복'));
  const hasIssue = isError || isWarning;

  // 셀 클래스 결정 (목업 v2 기준)
  function getClassName() {
    const classes = ['excel-cell'];

    // 오류/경고/중복 배경
    if (isError) classes.push('ce');
    else if (isDuplicate) classes.push('cd');
    else if (isWarning) classes.push('cw');

    // 선택 셀
    if (isSelected) classes.push('cs');

    // 헤더/레이블/숫자
    if (isHeader) classes.push('ch');
    if (isBold && !isHeader) classes.push('cb');
    if (isNumeric) classes.push('cn');

    return classes.join(' ');
  }

  // 셀 클릭: 셀 선택
  function handleClick() {
    const ref = toCellRef(row, col);
    setSelectedCell({ row, col, ref });

    if (window.__navigateToCell) {
      window.__navigateToCell(ref);
    }
  }

  // 호버 시 툴팁 표시
  function handleMouseEnter() {
    if (hasIssue && issues && issues.length > 0) {
      const rect = cellRef.current?.getBoundingClientRect();
      if (rect) {
        setAnchorRect(rect);
        setTooltipVisible(true);
      }
    }
  }

  function handleMouseLeave() {
    setTooltipVisible(false);
  }

  // 값 포맷팅
  function formatValue(val) {
    if (val === null || val === undefined) return '';
    if (typeof val === 'number') {
      return val.toLocaleString('ko-KR');
    }
    return String(val);
  }

  return (
    <td
      ref={cellRef}
      className={getClassName()}
      onClick={handleClick}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
      title={hasIssue && issues ? issues.map((i) => i.message).join('\n') : ''}
    >
      {formatValue(value)}
      {/* 이슈 툴팁 */}
      <CellTooltip
        issues={issues}
        anchorRect={anchorRect}
        visible={tooltipVisible}
      />
    </td>
  );
}
