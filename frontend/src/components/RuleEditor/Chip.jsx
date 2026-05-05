import { useState, useRef, useEffect } from 'react';

// =============================================
// 칩 컴포넌트
// 규칙 블록 내의 선택 가능한 조건 요소
// 타입: column(열), operator(연산자), value(값), action(동작)
// =============================================

// 타입별 색상 클래스
const TYPE_CLASS = {
  column: 'chip-col',
  operator: 'chip-op',
  value: 'chip-val',
  action: 'chip-act',
};

// 타입별 기본 옵션
const DEFAULT_OPTIONS = {
  column: ['열 선택...', 'A열', 'B열', 'C열', 'D열', 'E열', 'F열', 'G열'],
  operator: ['= (같음)', '≠ (같지 않음)', '> (초과)', '>= (이상)', '< (미만)', '<= (이하)', '포함', '미포함'],
  value: ['값 입력...', '0', '100', '1000', '10000'],
  action: ['오류 마킹', '경고 마킹', '정보 마킹', '셀 강조'],
};

export default function Chip({ type, value, options, onChange }) {
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const [inputMode, setInputMode] = useState(false);
  const [inputValue, setInputValue] = useState(value || '');
  const dropdownRef = useRef(null);
  const chipRef = useRef(null);

  const chipClass = TYPE_CLASS[type] || 'chip-col';
  const availableOptions = options || DEFAULT_OPTIONS[type] || [];

  // 드롭다운 외부 클릭 시 닫기
  useEffect(() => {
    function handleClickOutside(e) {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target) &&
          chipRef.current && !chipRef.current.contains(e.target)) {
        setDropdownOpen(false);
        setInputMode(false);
      }
    }
    if (dropdownOpen || inputMode) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [dropdownOpen, inputMode]);

  function handleChipClick() {
    setDropdownOpen(!dropdownOpen);
  }

  function handleOptionSelect(option) {
    if (option === '값 입력...') {
      setInputMode(true);
      setDropdownOpen(false);
      return;
    }
    onChange && onChange(option);
    setDropdownOpen(false);
  }

  function handleInputSubmit(e) {
    if (e.key === 'Enter') {
      onChange && onChange(inputValue);
      setInputMode(false);
    }
    if (e.key === 'Escape') {
      setInputMode(false);
    }
  }

  return (
    <div style={{ position: 'relative', display: 'inline-block' }}>
      {inputMode ? (
        // 직접 입력 모드
        <input
          autoFocus
          className={`chip ${chipClass}`}
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={handleInputSubmit}
          onBlur={() => {
            onChange && onChange(inputValue);
            setInputMode(false);
          }}
          style={{
            border: 'none',
            outline: 'none',
            minWidth: '80px',
            cursor: 'text',
          }}
        />
      ) : (
        // 일반 칩 버튼
        <span
          ref={chipRef}
          className={`chip ${chipClass}`}
          onClick={handleChipClick}
          title="클릭하여 변경"
        >
          {value || '선택...'}
        </span>
      )}

      {/* 드롭다운 옵션 목록 */}
      {dropdownOpen && (
        <div
          ref={dropdownRef}
          style={{
            position: 'absolute',
            top: '100%',
            left: 0,
            marginTop: 4,
            background: '#1c2128',
            border: '1px solid #30363d',
            borderRadius: 6,
            boxShadow: '0 4px 16px rgba(0,0,0,0.6)',
            zIndex: 1000,
            minWidth: '140px',
            maxHeight: '200px',
            overflowY: 'auto',
          }}
        >
          {availableOptions.map((opt) => (
            <div
              key={opt}
              onClick={() => handleOptionSelect(opt)}
              style={{
                padding: '7px 12px',
                fontSize: 11,
                color: '#c9d1d9',
                cursor: 'pointer',
                borderBottom: '1px solid #21262d',
              }}
              onMouseEnter={(e) => e.currentTarget.style.background = '#2d333b'}
              onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
            >
              {opt}
            </div>
          ))}
          {/* 직접 입력 옵션 (value 타입만) */}
          {type === 'value' && (
            <div
              onClick={() => handleOptionSelect('값 입력...')}
              style={{
                padding: '7px 12px',
                fontSize: 11,
                color: '#58a6ff',
                cursor: 'pointer',
              }}
              onMouseEnter={(e) => e.currentTarget.style.background = '#2d333b'}
              onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
            >
              직접 입력...
            </div>
          )}
        </div>
      )}
    </div>
  );
}
