import { useStore } from '../stores/store';

function escapeCell(value) {
  if (value == null) return '';
  return String(value).replace(/\r?\n/g, ' ').replace(/\|/g, '\\|');
}

function buildMarkdownTable(grid, minR, maxR, minC, maxC, hiddenSet) {
  const lines = [];
  const headerCells = [];
  for (let c = minC; c <= maxC; c++) {
    const v = grid[minR]?.[c];
    headerCells.push(escapeCell(v ?? ''));
  }
  lines.push(`| ${headerCells.join(' | ')} |`);
  lines.push(`| ${headerCells.map(() => '---').join(' | ')} |`);
  for (let r = minR + 1; r <= maxR; r++) {
    const rowCells = [];
    for (let c = minC; c <= maxC; c++) {
      if (hiddenSet && hiddenSet.has(`${r},${c}`)) {
        rowCells.push('');
        continue;
      }
      const v = grid[r]?.[c];
      rowCells.push(escapeCell(v ?? ''));
    }
    lines.push(`| ${rowCells.join(' | ')} |`);
  }
  return lines.join('\n');
}

function buildExcelMarkdown(cells, structure) {
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
      cellMap[sheetName][r][c] = cell.value;
    }
  });
  const sheetsMeta = structure?.sheets || [];
  const out = [];
  Object.keys(sheets).forEach((sheetName) => {
    out.push(`## 시트: ${sheetName}`);
    out.push('');
    const meta = sheetsMeta.find((s) => s.name === sheetName) || {};
    const mergeRanges = meta.merge_ranges || [];
    const hiddenSet = new Set();
    mergeRanges.forEach((m) => {
      for (let dr = 0; dr < m.rowspan; dr++) {
        for (let dc = 0; dc < m.colspan; dc++) {
          if (dr === 0 && dc === 0) continue;
          hiddenSet.add(`${m.row + dr},${m.col + dc}`);
        }
      }
    });
    const grid = cellMap[sheetName] || {};
    let minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;
    sheets[sheetName].forEach((cell) => {
      const n = cell.neighbors || {};
      if (n.row && n.col) {
        minRow = Math.min(minRow, n.row);
        maxRow = Math.max(maxRow, n.row);
        minCol = Math.min(minCol, n.col);
        maxCol = Math.max(maxCol, n.col);
      }
    });
    if (maxRow === 0) {
      sheets[sheetName].forEach((cell) => {
        out.push(`- ${cell.address}: ${cell.value ?? ''}`);
      });
      out.push('');
      return;
    }
    out.push(buildMarkdownTable(grid, minRow, maxRow, minCol, maxCol, hiddenSet));
    out.push('');
  });
  return out.join('\n');
}

function buildWordMarkdown(cells) {
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
  const out = [];
  paragraphs.forEach((p) => {
    out.push(p.value ?? '');
    out.push('');
  });
  Object.keys(tables).sort().forEach((tName) => {
    out.push(`## 표 ${tName}`);
    out.push('');
    const grid = {};
    let maxR = 0, maxC = 0;
    tables[tName].forEach(({ r, c, value }) => {
      if (!grid[r]) grid[r] = {};
      grid[r][c] = value;
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
    });
    out.push(buildMarkdownTable(grid, 0, maxR, 0, maxC, null));
    out.push('');
  });
  return out.join('\n');
}

function buildPdfMarkdown(cells) {
  return cells.map((c) => c.value ?? '').join('\n\n');
}

function buildImageMarkdown(cells) {
  const banner = '> ⚠ OCR 결과 — 일부 값이 누락될 수 있습니다';
  const hasGrid = cells.some((c) => c.address.startsWith('ocr_tbl:'));
  if (!hasGrid) {
    const lines = [banner, ''];
    cells.forEach((c) => lines.push(`- ${c.value ?? ''}`));
    return lines.join('\n');
  }
  const grid = {};
  let maxR = 0, maxC = 0;
  cells.forEach((cell) => {
    const m = cell.address.match(/R(\d+)C(\d+)/);
    if (m) {
      const r = parseInt(m[1]), c = parseInt(m[2]);
      if (!grid[r]) grid[r] = {};
      grid[r][c] = cell.value;
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
    }
  });
  return [banner, '', buildMarkdownTable(grid, 0, maxR, 0, maxC, null)].join('\n');
}

function buildOtherMarkdown(cells) {
  return cells.map((c) => `- ${c.address}: ${c.value ?? ''}`).join('\n');
}

export function buildMarkdown(parsedData, docName) {
  const lines = [`# ${docName}`, ''];
  if (!parsedData) {
    lines.push('_(데이터 없음)_');
    return lines.join('\n');
  }
  const fileType = parsedData.file_type;
  if (fileType) {
    lines.push(`_파일 형식: ${fileType}_`);
    lines.push('');
  }
  if (parsedData.metadata?.fallback === true) {
    lines.push(`⚠ 파싱 실패 — ${parsedData.metadata?.reason || 'unknown'}`);
    return lines.join('\n');
  }
  const cells = parsedData.cells || [];
  const structure = parsedData.structure || {};
  let body = '';
  switch (fileType) {
    case 'excel': body = buildExcelMarkdown(cells, structure); break;
    case 'word': body = buildWordMarkdown(cells); break;
    case 'pdf': body = buildPdfMarkdown(cells); break;
    case 'image': body = buildImageMarkdown(cells); break;
    default: body = buildOtherMarkdown(cells);
  }
  lines.push(body);
  return lines.join('\n');
}

export default function MarkdownView({ parsedData, docName }) {
  const selectedDocId = useStore((s) => s.selectedDocId);
  const storeParsed = useStore((s) => s.parsedData[selectedDocId]);
  const data = parsedData ?? storeParsed;
  if (!data) {
    return (
      <div style={{
        flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center',
        color: 'var(--text-sub)', fontSize: '13px',
      }}>파싱 데이터 로딩 중...</div>
    );
  }
  const markdown = buildMarkdown(data, docName || '문서');
  return (
    <pre style={{
      flex: 1, overflow: 'auto', padding: '16px 20px', margin: 0,
      fontFamily: 'ui-monospace, SFMono-Regular, Menlo, monospace',
      fontSize: '12px', lineHeight: '1.6', color: 'var(--text-main)',
      background: 'var(--bg-main)', whiteSpace: 'pre-wrap', wordBreak: 'break-word',
    }}>{markdown}</pre>
  );
}
