import { useState } from 'react';
import { useStore } from '../stores/store';

function ExcelGrid({ cells, structure }) {
  const [activeSheet, setActiveSheet] = useState(null);
  const sheets = {};
  const cellMap = {};
  cells.forEach((cell) => {
    const parts = cell.address.split('!');
    const sheetName = parts.length > 1 ? parts[0] : '_default';
    if (!sheets[sheetName]) sheets[sheetName] = [];
    sheets[sheetName].push(cell);
    const n = cell.neighbors || {};
    const r = n.row, c = n.col;
    if (r && c) {
      if (!cellMap[sheetName]) cellMap[sheetName] = {};
      if (!cellMap[sheetName][r]) cellMap[sheetName][r] = {};
      cellMap[sheetName][r][c] = cell;
    }
  });
  const sheetNames = Object.keys(sheets);
  const currentSheet = activeSheet || sheetNames[0] || '_default';
  const sheetCells = sheets[currentSheet] || [];

  // 시트 메타 (col_widths, merge_ranges 등)
  const sheetsMeta = structure?.sheets || [];
  const meta = sheetsMeta.find(s => s.name === currentSheet) || {};
  const colWidths = meta.col_widths || {};
  const mergeRanges = meta.merge_ranges || [];

  // 병합 hidden 셀 세트
  const hiddenSet = new Set();
  const mergeMap = {};
  mergeRanges.forEach(m => {
    for (let dr = 0; dr < m.rowspan; dr++) {
      for (let dc = 0; dc < m.colspan; dc++) {
        if (dr === 0 && dc === 0) continue;
        hiddenSet.add(`${m.row + dr},${m.col + dc}`);
      }
    }
    mergeMap[`${m.row},${m.col}`] = m;
  });

  // 그리드 범위
  let minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;
  sheetCells.forEach((cell) => {
    const n = cell.neighbors || {};
    if (n.row && n.col) {
      minRow = Math.min(minRow, n.row);
      maxRow = Math.max(maxRow, n.row);
      minCol = Math.min(minCol, n.col);
      maxCol = Math.max(maxCol, n.col);
    }
  });
  if (maxRow === 0) minRow = 1;

  const cm = cellMap[currentSheet] || {};

  return (
    <>
      {sheetNames.length > 1 && (
        <div style={{ display: 'flex', gap: '0', borderBottom: '1px solid var(--border)',
          background: 'var(--bg-panel)', flexShrink: 0 }}>
          {sheetNames.map((name) => (
            <button key={name} onClick={() => setActiveSheet(name)} style={{
              padding: '6px 16px', fontSize: '11px', cursor: 'pointer', border: 'none',
              borderBottom: name === currentSheet ? '2px solid var(--accent-blue)' : '2px solid transparent',
              background: 'transparent',
              color: name === currentSheet ? 'var(--text-main)' : 'var(--text-sub)',
              fontWeight: name === currentSheet ? 500 : 400,
            }}>{name}</button>
          ))}
        </div>
      )}
      {maxRow === 0 ? (
        <div style={{ color: 'var(--text-sub)', fontSize: '12px', textAlign: 'center', marginTop: '20px' }}>
          표시할 데이터가 없습니다
        </div>
      ) : (
        <table style={{ borderCollapse: 'collapse', fontSize: '11px', tableLayout: 'auto' }}>
          <tbody>
            {Array.from({ length: maxRow - minRow + 1 }, (_, ri) => ri + minRow).map((r) => (
              <tr key={r}>
                {Array.from({ length: maxCol - minCol + 1 }, (_, ci) => ci + minCol).map((c) => {
                  if (hiddenSet.has(`${r},${c}`)) return null;
                  const cell = cm[r]?.[c];
                  const n = cell?.neighbors || {};
                  const mg = mergeMap[`${r},${c}`];
                  const bgColor = n.bg_color;
                  const align = n.align === 'general' ? undefined : n.align;
                  const cw = colWidths[c];
                  return (
                    <td key={c}
                      rowSpan={mg?.rowspan > 1 ? mg.rowspan : undefined}
                      colSpan={mg?.colspan > 1 ? mg.colspan : undefined}
                      style={{
                        border: '1px solid var(--border)',
                        padding: '3px 6px',
                        color: 'var(--text-main)',
                        whiteSpace: 'nowrap',
                        background: bgColor || 'transparent',
                        textAlign: align,
                        minWidth: cw ? `${Math.round(cw * 7)}px` : undefined,
                        fontSize: '11px',
                      }}>{cell?.value ?? ''}</td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </>
  );
}

function WordView({ cells }) {
  const paragraphs = [];
  const tables = {};
  cells.forEach((cell) => {
    const addr = cell.address;
    if (addr.startsWith('para:')) {
      paragraphs.push(cell);
    } else {
      const tMatch = addr.match(/^(table\d+):R(\d+)C(\d+)/);
      if (tMatch) {
        const tName = tMatch[1];
        if (!tables[tName]) tables[tName] = [];
        tables[tName].push({ r: parseInt(tMatch[2]), c: parseInt(tMatch[3]), value: cell.value });
      }
    }
  });
  const tableNames = Object.keys(tables).sort();
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
      {paragraphs.length > 0 && (
        <div>
          <div style={{ fontSize: '11px', color: 'var(--text-sub)', marginBottom: '6px', fontWeight: 500 }}>
            문단 ({paragraphs.length})
          </div>
          {paragraphs.map((p, i) => (
            <div key={i} style={{ fontSize: '12px', color: 'var(--text-main)', padding: '2px 0',
              lineHeight: '1.6', borderBottom: '1px solid rgba(255,255,255,0.04)' }}>
              {p.value}
            </div>
          ))}
        </div>
      )}
      {tableNames.map((tName) => {
        const tCells = tables[tName];
        const grid = {};
        let maxR = 0, maxC = 0;
        tCells.forEach(({ r, c, value }) => {
          if (!grid[r]) grid[r] = {};
          grid[r][c] = value;
          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, c);
        });
        return (
          <div key={tName}>
            <div style={{ fontSize: '11px', color: 'var(--text-sub)', marginBottom: '6px', fontWeight: 500 }}>
              {tName}
            </div>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px', marginBottom: '8px' }}>
              <tbody>
                {Array.from({ length: maxR + 1 }, (_, ri) => ri).map((r) => (
                  <tr key={r}>
                    {Array.from({ length: maxC + 1 }, (_, ci) => ci).map((c) => (
                      <td key={c} style={{
                        border: '1px solid var(--border)', padding: '4px 6px',
                        color: 'var(--text-main)', whiteSpace: 'pre-wrap',
                        background: r === 0 ? 'var(--bg-panel)' : 'transparent',
                        fontWeight: r === 0 ? 500 : 400,
                      }}>{grid[r]?.[c] ?? ''}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
      })}
    </div>
  );
}

function TextListView({ cells, label }) {
  if (!cells || cells.length === 0) {
    return (
      <div style={{ color: 'var(--text-sub)', fontSize: '12px', textAlign: 'center', marginTop: '20px' }}>
        표시할 데이터가 없습니다
      </div>
    );
  }
  return (
    <div>
      <div style={{ fontSize: '11px', color: 'var(--text-sub)', marginBottom: '6px', fontWeight: 500 }}>
        {label} ({cells.length}개 항목)
      </div>
      {cells.map((cell, i) => (
        <div key={i} style={{
          fontSize: '12px', color: 'var(--text-main)', padding: '3px 0',
          borderBottom: '1px solid rgba(255,255,255,0.04)',
        }}>
          <span style={{ fontSize: '10px', color: 'var(--text-sub)', marginRight: '8px' }}>{cell.address}</span>
          {cell.value}
        </div>
      ))}
    </div>
  );
}

function OcrGridView({ cells }) {
  // ocr_tbl:R{row}C{col} 형식이면 그리드, 아니면 리스트
  const hasGrid = cells.some(c => c.address.startsWith('ocr_tbl:'));
  if (!hasGrid) return <TextListView cells={cells} label="OCR 결과" />;

  const grid = {};
  let maxR = 0, maxC = 0;
  cells.forEach(cell => {
    const m = cell.address.match(/R(\d+)C(\d+)/);
    if (m) {
      const r = parseInt(m[1]), c = parseInt(m[2]);
      if (!grid[r]) grid[r] = {};
      grid[r][c] = cell.value;
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
    }
  });

  return (
    <div>
      <div style={{ fontSize: '11px', color: 'var(--text-sub)', marginBottom: '6px', fontWeight: 500 }}>
        OCR 결과 ({maxR + 1}행 × {maxC + 1}열)
      </div>
      <table style={{ borderCollapse: 'collapse', fontSize: '11px', width: '100%' }}>
        <tbody>
          {Array.from({ length: maxR + 1 }, (_, r) => (
            <tr key={r}>
              {Array.from({ length: maxC + 1 }, (_, c) => (
                <td key={c} style={{
                  border: '1px solid var(--border)', padding: '3px 6px',
                  color: 'var(--text-main)', whiteSpace: 'nowrap',
                  background: r === 0 ? 'var(--bg-panel)' : 'transparent',
                  fontWeight: r === 0 ? 500 : 400,
                }}>{grid[r]?.[c] ?? ''}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function DataTable() {
  const selectedDocId = useStore((s) => s.selectedDocId);
  const parsedData = useStore((s) => s.parsedData[selectedDocId]);
  if (!parsedData) {
    return (
      <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center',
        color: 'var(--text-sub)', fontSize: '13px' }}>
        파싱 데이터 로딩 중...
      </div>
    );
  }

  const cells = parsedData.cells || [];
  const fileType = parsedData.file_type;
  const structure = parsedData.structure || {};
  const hasFallback = parsedData.metadata?.fallback === true;
  const fallbackReason = parsedData.metadata?.reason || '';

  if (hasFallback) {
    const reasons = {
      'tesseract_not_installed': 'pytesseract 미설치 — Windows OCR도 실패',
      'windows_ocr_no_result': 'Windows OCR 결과 없음 — 이미지 품질 확인 필요',
      'file_open_error': '파일을 열 수 없습니다',
      'ocr_error': 'OCR 처리 실패',
      'pypdf_not_installed': 'pypdf 미설치 — pip install pypdf 필요',
    };
    return (
      <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center',
        flexDirection: 'column', gap: '8px' }}>
        <span style={{ fontSize: '13px', color: 'var(--color-orange)' }}>
          ⚠ {reasons[fallbackReason] || '파싱 실패'}
        </span>
        <span style={{ fontSize: '11px', color: 'var(--text-sub)' }}>
          {fileType === 'pdf' && 'Acrobat Pro COM 연결 불가 — PDF 파일 업로드 또는 Acrobat 설치 필요'}
          {fileType === 'image' && 'pip install pytesseract 후 Tesseract OCR 바이너리 설치 필요'}
        </span>
      </div>
    );
  }

  return (
    <div style={{ padding: '8px' }}>
        {fileType === 'excel' && <ExcelGrid cells={cells} structure={structure} />}
        {fileType === 'word' && <WordView cells={cells} />}
        {fileType === 'pdf' && <TextListView cells={cells} label="PDF 텍스트" />}
        {fileType === 'image' && <OcrGridView cells={cells} />}
        {!['excel', 'word', 'pdf', 'image'].includes(fileType) && <ExcelGrid cells={cells} structure={structure} />}
    </div>
  );
}
