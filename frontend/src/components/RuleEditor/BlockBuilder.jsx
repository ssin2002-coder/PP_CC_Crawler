import RuleBlock from './RuleBlock.jsx';

// =============================================
// 블록 빌더 컴포넌트
// 규칙 블록 목록 관리 (추가/수정/삭제)
// =============================================

// 새 블록 추가 시 기본 구조
function createNewBlock(type) {
  return {
    id: `block_${Date.now()}`,
    type,
    chips:
      type === 'then'
        ? [{ type: 'action', value: '동작 선택...' }]
        : [
            { type: 'column', value: '열 선택...' },
            { type: 'operator', value: '조건 선택...' },
            { type: 'value', value: '값 입력...' },
          ],
  };
}

export default function BlockBuilder({ blocks, onChange }) {
  // 블록 수정
  function handleBlockUpdate(updatedBlock) {
    const newBlocks = blocks.map((b) =>
      b.id === updatedBlock.id ? updatedBlock : b
    );
    onChange && onChange(newBlocks);
  }

  // 블록 삭제 (IF 블록은 첫 번째라면 삭제 불가)
  function handleBlockDelete(blockId) {
    const block = blocks.find((b) => b.id === blockId);
    if (block && block.type === 'if' && blocks.filter((b) => b.type === 'if').length <= 1) {
      alert('최소 하나의 IF 블록이 필요합니다.');
      return;
    }
    const newBlocks = blocks.filter((b) => b.id !== blockId);
    onChange && onChange(newBlocks);
  }

  // AND 블록 추가 (THEN 블록 앞에 삽입)
  function handleAddAnd() {
    const thenIndex = blocks.findIndex((b) => b.type === 'then');
    const newBlock = createNewBlock('and');
    const newBlocks = [...blocks];
    if (thenIndex >= 0) {
      newBlocks.splice(thenIndex, 0, newBlock);
    } else {
      newBlocks.push(newBlock);
    }
    onChange && onChange(newBlocks);
  }

  // OR 블록 추가
  function handleAddOr() {
    const thenIndex = blocks.findIndex((b) => b.type === 'then');
    const newBlock = createNewBlock('or');
    const newBlocks = [...blocks];
    if (thenIndex >= 0) {
      newBlocks.splice(thenIndex, 0, newBlock);
    } else {
      newBlocks.push(newBlock);
    }
    onChange && onChange(newBlocks);
  }

  return (
    <div className="block-area">
      {/* 블록 렌더링 */}
      {blocks.map((block) => (
        <RuleBlock
          key={block.id}
          block={block}
          onUpdate={handleBlockUpdate}
          onDelete={handleBlockDelete}
        />
      ))}

      {/* 블록 추가 버튼 */}
      <div style={{ display: 'flex', gap: 8 }}>
        <div className="add-block-btn" onClick={handleAddAnd} style={{ flex: 1 }}>
          + AND 조건 추가
        </div>
        <div className="add-block-btn" onClick={handleAddOr} style={{ flex: 1 }}>
          + OR 조건 추가
        </div>
      </div>
    </div>
  );
}
