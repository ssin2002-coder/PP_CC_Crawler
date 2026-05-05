import Chip from './Chip.jsx';

// =============================================
// 규칙 블록 컴포넌트
// IF / AND / THEN 블록을 렌더링
// =============================================

// 블록 타입별 스타일 설정
const BLOCK_CONFIG = {
  if: {
    className: 'rule-block if-block',
    labelClass: 'block-label label-if',
    labelText: 'IF',
  },
  and: {
    className: 'rule-block and-block',
    labelClass: 'block-label label-and',
    labelText: 'AND',
  },
  or: {
    className: 'rule-block and-block', // OR는 AND와 같은 스타일
    labelClass: 'block-label label-and',
    labelText: 'OR',
  },
  then: {
    className: 'rule-block then-block',
    labelClass: 'block-label label-then',
    labelText: 'THEN',
  },
};

export default function RuleBlock({ block, onUpdate, onDelete }) {
  const config = BLOCK_CONFIG[block.type] || BLOCK_CONFIG.if;

  // 개별 칩 값 변경
  function handleChipChange(chipIndex, newValue) {
    const updatedChips = block.chips.map((chip, idx) =>
      idx === chipIndex ? { ...chip, value: newValue } : chip
    );
    onUpdate && onUpdate({ ...block, chips: updatedChips });
  }

  return (
    <div className={config.className} style={{ position: 'relative' }}>
      {/* 삭제 버튼 (호버 시 표시 - CSS로 처리) */}
      <button
        className="block-delete-btn"
        onClick={() => onDelete && onDelete(block.id)}
        title="블록 삭제"
        style={{
          position: 'absolute',
          top: 8,
          right: 8,
          background: 'transparent',
          border: 'none',
          color: '#8b949e',
          cursor: 'pointer',
          fontSize: 12,
          padding: '2px 4px',
          borderRadius: 3,
          lineHeight: 1,
        }}
        onMouseEnter={(e) => e.currentTarget.style.color = '#f85149'}
        onMouseLeave={(e) => e.currentTarget.style.color = '#8b949e'}
      >
        ✕
      </button>

      {/* 블록 레이블 (IF / AND / THEN) */}
      <span className={config.labelClass}>{config.labelText}</span>

      {/* 칩 목록 */}
      <div className="block-chips">
        {block.chips.map((chip, idx) => (
          <Chip
            key={idx}
            type={chip.type}
            value={chip.value}
            onChange={(newValue) => handleChipChange(idx, newValue)}
          />
        ))}
      </div>
    </div>
  );
}
