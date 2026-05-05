import useStore from '../../stores/store.js';
import { fetchData } from '../../utils/api.js';

// =============================================
// ExcelToolbar 컴포넌트
// sheet-strip 스타일: 시트 탭 + 우측 정보 텍스트
// =============================================

export default function ExcelToolbar() {
  const excelData = useStore((s) => s.excelData);
  const setExcelData = useStore((s) => s.setExcelData);
  const setActiveSheet = useStore((s) => s.setActiveSheet);

  const sheets = excelData.sheets || [];
  const activeSheet = excelData.activeSheet;
  const cells = excelData.cells || [];

  // 행/열 수 계산
  let rowCount = 0;
  let colCount = 0;
  if (cells.length > 0) {
    if (Array.isArray(cells[0])) {
      rowCount = cells.length;
      colCount = Math.max(...cells.map((r) => r.length), 0);
    } else {
      cells.forEach((c) => {
        if ((c.row || 0) > rowCount) rowCount = c.row || 0;
        if ((c.col || 0) > colCount) colCount = c.col || 0;
      });
    }
  }

  const infoText = rowCount > 0
    ? `${rowCount} rows × ${colCount} cols`
    : '';

  const handleSheetClick = async (sheetName) => {
    if (sheetName === activeSheet) return;
    const wb = excelData.activeWorkbook;
    if (!wb) return;
    const wbName = typeof wb === 'string' ? wb : wb.name;

    setActiveSheet(sheetName);
    const { data } = await fetchData(wbName, sheetName);
    if (data && data.cells) {
      setExcelData({ cells: data.cells });
    }
  };

  return (
    <div className="sheet-strip">
      {sheets.length === 0 ? (
        <div className="st active">시트 없음</div>
      ) : (
        sheets.map((sheet) => (
          <div
            key={sheet}
            className={`st${activeSheet === sheet ? ' active' : ''}`}
            onClick={() => handleSheetClick(sheet)}
            title={sheet}
          >
            {sheet}
          </div>
        ))
      )}

      {infoText && (
        <span className="info">{infoText}</span>
      )}
    </div>
  );
}
