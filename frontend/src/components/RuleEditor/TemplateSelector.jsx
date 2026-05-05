// =============================================
// 템플릿 선택기 컴포넌트
// 규칙 편집기에서 사전 정의된 템플릿 선택
// =============================================

// 6개 템플릿 정의
const TEMPLATES = [
  {
    id: 'sum_check',
    label: '합계 검증',
    defaultBlocks: [
      {
        id: 'b1',
        type: 'if',
        chips: [
          { type: 'column', value: '합산 대상 열' },
          { type: 'operator', value: '≠ (같지 않음)' },
          { type: 'column', value: '합계 셀' },
        ],
      },
      {
        id: 'b2',
        type: 'then',
        chips: [
          { type: 'action', value: '오류 마킹' },
          { type: 'value', value: '"합계 불일치"' },
        ],
      },
    ],
  },
  {
    id: 'outlier',
    label: '이상치 탐지',
    defaultBlocks: [
      {
        id: 'b1',
        type: 'if',
        chips: [
          { type: 'column', value: '대상 열' },
          { type: 'operator', value: '> (초과)' },
          { type: 'value', value: '평균 + 3σ' },
        ],
      },
      {
        id: 'b2',
        type: 'then',
        chips: [
          { type: 'action', value: '경고 마킹' },
          { type: 'value', value: '"이상치 탐지"' },
        ],
      },
    ],
  },
  {
    id: 'duplicate',
    label: '중복 검출',
    defaultBlocks: [
      {
        id: 'b1',
        type: 'if',
        chips: [
          { type: 'column', value: '검사 열' },
          { type: 'operator', value: '= (중복)' },
          { type: 'value', value: '이전 행' },
        ],
      },
      {
        id: 'b2',
        type: 'then',
        chips: [
          { type: 'action', value: '경고 마킹' },
          { type: 'value', value: '"중복 항목"' },
        ],
      },
    ],
  },
  {
    id: 'range_exceed',
    label: '범위 초과',
    defaultBlocks: [
      {
        id: 'b1',
        type: 'if',
        chips: [
          { type: 'column', value: '대상 열' },
          { type: 'operator', value: '> (초과)' },
          { type: 'value', value: '최대값' },
        ],
      },
      {
        id: 'b2',
        type: 'then',
        chips: [
          { type: 'action', value: '오류 마킹' },
          { type: 'value', value: '"범위 초과"' },
        ],
      },
    ],
  },
  {
    id: 'cross_check',
    label: '교차 검증',
    defaultBlocks: [
      {
        id: 'b1',
        type: 'if',
        chips: [
          { type: 'column', value: '열 A' },
          { type: 'operator', value: '+ (합산)' },
          { type: 'column', value: '열 B' },
          { type: 'operator', value: '≠ (같지 않음)' },
          { type: 'column', value: '총합계 열' },
        ],
      },
      {
        id: 'b2',
        type: 'then',
        chips: [
          { type: 'action', value: '오류 마킹' },
          { type: 'value', value: '"교차 검증 실패"' },
        ],
      },
    ],
  },
  {
    id: 'custom',
    label: '사용자 정의',
    defaultBlocks: [
      {
        id: 'b1',
        type: 'if',
        chips: [
          { type: 'column', value: '열 선택...' },
          { type: 'operator', value: '조건 선택...' },
          { type: 'value', value: '값 입력...' },
        ],
      },
      {
        id: 'b2',
        type: 'then',
        chips: [
          { type: 'action', value: '동작 선택...' },
        ],
      },
    ],
  },
];

export default function TemplateSelector({ selectedTemplate, onSelect }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ fontSize: 11, color: '#8b949e', marginBottom: 8 }}>
        템플릿 선택:
      </div>
      <div className="template-selector">
        {TEMPLATES.map((tmpl) => (
          <span
            key={tmpl.id}
            className={`template-chip${selectedTemplate === tmpl.id ? ' selected' : ''}`}
            onClick={() => onSelect(tmpl.id, tmpl.defaultBlocks)}
          >
            {tmpl.label}
          </span>
        ))}
      </div>
    </div>
  );
}

// 템플릿 기본값 내보내기 (외부에서 사용 가능)
export { TEMPLATES };
