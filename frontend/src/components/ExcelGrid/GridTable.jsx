import { useRef, useMemo, useCallback } from 'react';
import { useVirtualizer } from '@tanstack/react-virtual';
import GridCell from './GridCell.jsx';
import useStore from '../../stores/store.js';
import { colToLetter, toCellRef } from '../../utils/cellRef.js';
import { navigateToCell } from '../../utils/api.js';

// =============================================
// 그리드 테이블 컴포넌트
// 가상 스크롤링을 활용한 Excel 셀 렌더링
// =============================================

// 기본 열 너비 및 행 높이
const DEFAULT_COL_WIDTH = 100;
const ROW_HEADER_WIDTH = 36;
const ROW_HEIGHT = 24;
const HEADER_HEIGHT = 26;

export default function GridTable() {
  const scrollContainerRef = useRef(null);

  const cells = useStore((s) => s.excelData.cells);
  const issues = useStore((s) => s.issues);
  const selectedCell = useStore((s) => s.selectedCell);
  const activeRule = useStore((s) => s.activeRule);
  const getRelatedIssues = useStore((s) => s.getRelatedIssues);
  const getRelatedCells = useStore((s) => s.getRelatedCells);
  const excelData = useStore((s) => s.excelData);

  // 셀 데이터를 2D 맵으로 변환 {rowIdx: {colIdx: cellData}}
  // 백엔드 cells는 2D 배열 [[cell, ...], ...], 각 cell의 row/col은 1-based
  const cellMap = useMemo(() => {
    const map = {};
    if (!cells || cells.length === 0) return map;
    // 2D 배열 형태 (cells[rowIdx][colIdx])
    if (Array.isArray(cells[0])) {
      cells.forEach((row, ri) => {
        row.forEach((cell, ci) => {
          if (!map[ri]) map[ri] = {};
          map[ri][ci] = cell;
        });
      });
    } else {
      // flat 배열 폴백 (row/col 1-based → 0-based)
      cells.forEach((cell) => {
        const r = (cell.row || 1) - 1;
        const c = (cell.col || 1) - 1;
        if (!map[r]) map[r] = {};
        map[r][c] = cell;
      });
    }
    return map;
  }, [cells]);

  // 행/열 범위 계산
  const { maxRow, maxCol } = useMemo(() => {
    if (!cells || cells.length === 0) return { maxRow: 50, maxCol: 10 };
    // 2D 배열
    if (Array.isArray(cells[0])) {
      const mr = cells.length;
      const mc = Math.max(...cells.map((r) => r.length));
      return { maxRow: Math.max(mr + 5, 50), maxCol: Math.max(mc + 3, 10) };
    }
    // flat 배열 폴백
    let mr = 0, mc = 0;
    cells.forEach((c) => {
      if (c.row > mr) mr = c.row;
      if (c.col > mc) mc = c.col;
    });
    return { maxRow: Math.max(mr + 5, 50), maxCol: Math.max(mc + 3, 10) };
  }, [cells]);

  // 이슈 맵: {cellRef: [issue, ...]} (snake_case 대응)
  const issueMap = useMemo(() => {
    const map = {};
    issues.forEach((issue) => {
      const ref = issue.cell_ref || issue.cellRef;
      if (!ref) return;
      if (!map[ref]) map[ref] = [];
      map[ref].push(issue);
    });
    return map;
  }, [issues]);

  // 활성 규칙과 관련된 셀 목록
  const relatedCells = useMemo(
    () => new Set(getRelatedCells(activeRule)),
    [activeRule, getRelatedCells]
  );

  // 가상 스크롤: 행
  const rowVirtualizer = useVirtualizer({
    count: maxRow,
    getScrollElement: () => scrollContainerRef.current,
    estimateSize: () => ROW_HEIGHT,
    overscan: 5,
  });

  // 가상 스크롤: 열
  const colVirtualizer = useVirtualizer({
    count: maxCol,
    horizontal: true,
    getScrollElement: () => scrollContainerRef.current,
    estimateSize: () => DEFAULT_COL_WIDTH,
    overscan: 3,
  });

  const virtualRows = rowVirtualizer.getVirtualItems();
  const virtualCols = colVirtualizer.getVirtualItems();

  // 셀 클릭 시 Excel 파일로 네비게이션
  const handleNavigate = useCallback(async (cellRef) => {
    const wb = excelData.activeWorkbook;
    if (wb && excelData.activeSheet) {
      const wbName = typeof wb === 'string' ? wb : wb.name;
      await navigateToCell(wbName, excelData.activeSheet, cellRef);
    }
  }, [excelData.activeWorkbook, excelData.activeSheet]);

  // 전역 navigate 함수 등록 (GridCell에서 호출)
  window.__navigateToCell = handleNavigate;

  const totalWidth = ROW_HEADER_WIDTH + virtualCols.reduce((sum, vc) => sum + vc.size, 0);

  return (
    <div
      ref={scrollContainerRef}
      className="grid-scroll"
      style={{ position: 'relative' }}
    >
      <table
        className="g"
        style={{
          width: `${totalWidth}px`,
          tableLayout: 'fixed',
        }}
      >
        {/* 열 너비 정의 */}
        <colgroup>
          <col style={{ width: ROW_HEADER_WIDTH }} />
          {virtualCols.map((vc) => (
            <col key={vc.index} style={{ width: vc.size }} />
          ))}
        </colgroup>

        {/* 헤더 행 (A, B, C, ...) - 상단 고정 */}
        <thead>
          <tr style={{ height: HEADER_HEIGHT }}>
            {/* 빈 코너 셀 */}
            <th className="corner" style={{ position: 'sticky', left: 0, top: 0, zIndex: 3 }} />
            {virtualCols.map((vc) => (
              <th
                key={vc.index}
                style={{
                  position: 'sticky',
                  top: 0,
                  zIndex: 2,
                  minWidth: DEFAULT_COL_WIDTH,
                  textAlign: 'center',
                }}
              >
                {colToLetter(vc.index)}
              </th>
            ))}
          </tr>
        </thead>

        {/* 데이터 행 영역 */}
        <tbody>
          {/* 가상 스크롤 상단 여백 */}
          {virtualRows.length > 0 && virtualRows[0].start > 0 && (
            <tr style={{ height: virtualRows[0].start }}>
              <td colSpan={virtualCols.length + 1} />
            </tr>
          )}

          {virtualRows.map((vr) => {
            const rowIndex = vr.index;
            const rowData = cellMap[rowIndex] || {};

            return (
              <tr key={rowIndex} style={{ height: ROW_HEIGHT }}>
                {/* 행 번호 셀 - 좌측 고정 */}
                <td
                  className="rn"
                  style={{ position: 'sticky', left: 0, zIndex: 1 }}
                >
                  {rowIndex + 1}
                </td>

                {/* 가상 스크롤 좌측 여백 */}
                {virtualCols.length > 0 && virtualCols[0].start > ROW_HEADER_WIDTH && (
                  <td
                    style={{ width: virtualCols[0].start - ROW_HEADER_WIDTH }}
                  />
                )}

                {/* 데이터 셀 */}
                {virtualCols.map((vc) => {
                  const colIndex = vc.index;
                  const cell = rowData[colIndex] || null;
                  const cellRef = toCellRef(rowIndex, colIndex);
                  const cellIssues = issueMap[cellRef] || [];
                  const isSelected =
                    selectedCell &&
                    selectedCell.row === rowIndex &&
                    selectedCell.col === colIndex;
                  const isError = cellIssues.some((i) => i.severity === 'error');
                  const isWarning =
                    !isError && cellIssues.some((i) => i.severity === 'warning' || i.severity === 'warn');
                  const isRelated = relatedCells.has(cellRef);

                  return (
                    <GridCell
                      key={colIndex}
                      row={rowIndex}
                      col={colIndex}
                      cell={cell}
                      isSelected={isSelected}
                      isError={isError}
                      isWarning={isWarning}
                      isRelated={isRelated}
                      issues={cellIssues}
                    />
                  );
                })}

                {/* 가상 스크롤 우측 여백 */}
                {virtualCols.length > 0 && (() => {
                  const lastVC = virtualCols[virtualCols.length - 1];
                  const rightPad = rowVirtualizer.getTotalSize() - lastVC.end;
                  return rightPad > 0 ? <td style={{ width: rightPad }} /> : null;
                })()}
              </tr>
            );
          })}

          {/* 가상 스크롤 하단 여백 */}
          {virtualRows.length > 0 && (() => {
            const lastVR = virtualRows[virtualRows.length - 1];
            const bottomPad = rowVirtualizer.getTotalSize() - lastVR.end;
            return bottomPad > 0 ? (
              <tr style={{ height: bottomPad }}>
                <td colSpan={virtualCols.length + 1} />
              </tr>
            ) : null;
          })()}
        </tbody>
      </table>
    </div>
  );
}
