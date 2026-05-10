import { useState } from 'react';
import { useStore } from '../stores/store';
export default function DataTable() {
  const selectedDocId = useStore((s) => s.selectedDocId);
  const parsedData = useStore((s) => s.parsedData[selectedDocId]);
  const [activeSheet, setActiveSheet] = useState(null);
  if (!parsedData) {
    return (
      <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center',
        color: 'var(--text-sub)', fontSize: '13px' }}>
        파싱 데이터 로딩 중...
      </div>
    );
  }
  const cells = parsedData.cells || [];
  const sheets = {};
  cells.forEach((cell) => {
    const parts = cell.address.split('!');
    const sheetName = parts.length > 1 ? parts[0] : '_default';
    if (!sheets[sheetName]) sheets[sheetName] = [];
    sheets[sheetName].push(cell);
  });
  const sheetNames = Object.keys(sheets);
  const currentSheet = activeSheet || sheetNames[0] || '_default';
  const sheetCells = sheets[currentSheet] || [];
  const grid = {};
  let maxRow = 0, maxCol = 0;
  sheetCells.forEach((cell) => {
    const match = cell.address.match(/R(\d+)C(\d+)/);
    if (match) {
      const r = parseInt(match[1]), c = parseInt(match[2]);
      if (!grid[r]) grid[r] = {};
      grid[r][c] = cell.value;
      maxRow = Math.max(maxRow, r);
      maxCol = Math.max(maxCol, c);
    }
  });
  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
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
      <div style={{ flex: 1, overflow: 'auto', padding: '8px' }}>
        {maxRow === 0 ? (
          <div style={{ color: 'var(--text-sub)', fontSize: '12px', textAlign: 'center', marginTop: '20px' }}>
            표시할 데이터가 없습니다
          </div>
        ) : (
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px' }}>
            <tbody>
              {Array.from({ length: maxRow }, (_, ri) => ri + 1).map((r) => (
                <tr key={r}>
                  {Array.from({ length: maxCol }, (_, ci) => ci + 1).map((c) => (
                    <td key={c} style={{
                      border: '1px solid var(--border)', padding: '4px 6px',
                      color: 'var(--text-main)', whiteSpace: 'nowrap',
                      background: r === 1 ? 'var(--bg-panel)' : 'transparent',
                      fontWeight: r === 1 ? 500 : 400,
                    }}>{grid[r]?.[c] ?? ''}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}
