import ExcelToolbar from './ExcelToolbar.jsx';
import GridTable from './GridTable.jsx';

// =============================================
// ExcelArea (GridArea) 컨테이너
// sheet-strip + grid-scroll + float-info 구조
// =============================================

export default function ExcelArea() {
  return (
    <div className="grid-wrap">
      {/* 시트 탭 스트립 */}
      <ExcelToolbar />

      {/* Excel 셀 그리드 */}
      <GridTable />

      {/* 플로팅 정보 칩 */}
      <div className="float-info">
        <div className="fc">클릭 시 Excel 셀 이동</div>
        <div className="fc"><kbd>+</kbd> 규칙 추가</div>
      </div>
    </div>
  );
}
