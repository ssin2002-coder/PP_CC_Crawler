import useStore from '../stores/store.js';

// =============================================
// HeroStrip 컴포넌트
// 110px 통계 띠 - Errors / Warnings / Passed
// countUp 애니메이션 포함
// =============================================

export default function HeroStrip() {
  const summary = useStore((s) => s.summary);

  const errors = summary.errors || 0;
  const warnings = summary.warnings || 0;
  // ok는 boolean이므로 총 셀 수에서 이슈 수를 빼서 계산
  const cells = useStore((s) => s.excelData.cells);
  const totalCells = Array.isArray(cells) && cells.length > 0
    ? (Array.isArray(cells[0]) ? cells.length * (cells[0].length || 1) : cells.length)
    : 0;
  const passed = Math.max(totalCells - errors - warnings, 0);

  return (
    <div className="hero">
      <div className="hero-stat">
        <div className="hero-num red">{errors}</div>
        <div className="hero-label">Errors</div>
      </div>
      <div className="hero-stat">
        <div className="hero-num amber">{warnings}</div>
        <div className="hero-label">Warnings</div>
      </div>
      <div className="hero-stat">
        <div className="hero-num green">{passed}</div>
        <div className="hero-label">Passed</div>
      </div>
    </div>
  );
}
