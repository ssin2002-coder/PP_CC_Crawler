import { useState, useEffect } from 'react';
import useStore from '../../stores/store.js';
import { useRules } from '../../hooks/useRules.js';
import TemplateSelector, { TEMPLATES } from './TemplateSelector.jsx';
import BlockBuilder from './BlockBuilder.jsx';

// =============================================
// 규칙 편집기 모달 컴포넌트
// 신규 규칙 생성 및 기존 규칙 편집
// =============================================

const SEVERITY_OPTIONS = [
  { id: 'error', label: '오류 (ERROR)', color: '#f85149', borderColor: '#3d1a1a' },
  { id: 'warning', label: '경고 (WARN)', color: '#d29922', borderColor: '#3d2e00' },
  { id: 'info', label: '정보 (INFO)', color: '#58a6ff', borderColor: '#0c2d6b' },
];

// 신규 규칙 기본값
const DEFAULT_RULE = {
  name: '',
  description: '',
  type: 'manual',
  severity: 'error',
  enabled: true,
  blocks: [
    {
      id: 'b_if',
      type: 'if',
      chips: [
        { type: 'column', value: '열 선택...' },
        { type: 'operator', value: '조건 선택...' },
        { type: 'value', value: '값 입력...' },
      ],
    },
    {
      id: 'b_then',
      type: 'then',
      chips: [
        { type: 'action', value: '동작 선택...' },
        { type: 'value', value: '"오류 메시지"' },
      ],
    },
  ],
};

export default function RuleEditorModal() {
  const isOpen = useStore((s) => s.isRuleEditorOpen);
  const editingRule = useStore((s) => s.editingRule);
  const closeRuleEditor = useStore((s) => s.closeRuleEditor);

  const { createRule, updateRule, deleteRule } = useRules();

  // 편집 폼 상태
  const [formData, setFormData] = useState(DEFAULT_RULE);
  const [selectedTemplate, setSelectedTemplate] = useState(null);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState(null);

  // 편집 대상 규칙이 바뀔 때 폼 초기화
  useEffect(() => {
    if (editingRule) {
      setFormData({ ...DEFAULT_RULE, ...editingRule });
      setSelectedTemplate(editingRule.template || null);
    } else {
      setFormData(DEFAULT_RULE);
      setSelectedTemplate(null);
    }
    setError(null);
  }, [editingRule, isOpen]);

  if (!isOpen) return null;

  // 템플릿 선택 시 블록 업데이트
  function handleTemplateSelect(templateId, defaultBlocks) {
    setSelectedTemplate(templateId);
    setFormData((prev) => ({
      ...prev,
      template: templateId,
      blocks: defaultBlocks || prev.blocks,
    }));
  }

  // 폼 필드 변경
  function handleFieldChange(field, value) {
    setFormData((prev) => ({ ...prev, [field]: value }));
  }

  // 저장
  async function handleSave() {
    if (!formData.name.trim()) {
      setError('규칙 이름을 입력하십시오.');
      return;
    }
    setIsSaving(true);
    setError(null);

    const ruleData = { ...formData, template: selectedTemplate };

    let result;
    if (editingRule && editingRule.id) {
      result = await updateRule(editingRule.id, ruleData);
    } else {
      result = await createRule(ruleData);
    }

    setIsSaving(false);

    if (result.error) {
      setError(result.error);
    } else {
      closeRuleEditor();
    }
  }

  // 규칙 삭제
  async function handleDelete() {
    if (!editingRule || !editingRule.id) return;
    if (!window.confirm(`"${editingRule.name}" 규칙을 삭제하시겠습니까?`)) return;

    setIsSaving(true);
    const result = await deleteRule(editingRule.id);
    setIsSaving(false);

    if (result.error) {
      setError(result.error);
    } else {
      closeRuleEditor();
    }
  }

  // 테스트 실행 (현재는 콘솔 출력)
  function handleTest() {
    console.log('[테스트] 규칙 데이터:', formData);
    alert('테스트 실행 기능은 백엔드 연동 후 활성화됩니다.');
  }

  // 오버레이 클릭 시 닫기
  function handleOverlayClick(e) {
    if (e.target === e.currentTarget) {
      closeRuleEditor();
    }
  }

  const isEditing = !!(editingRule && editingRule.id);

  return (
    <div className="modal-overlay" onClick={handleOverlayClick}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        {/* 모달 헤더 */}
        <div className="modal-header">
          <span className="modal-title">
            {isEditing ? `규칙 편집: ${editingRule.name}` : '새 규칙 추가'}
          </span>
          <span className="modal-close" onClick={closeRuleEditor} title="닫기">
            &times;
          </span>
        </div>

        {/* 모달 본문 */}
        <div className="modal-body">
          {/* 규칙 이름 입력 */}
          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 11, color: '#8b949e', display: 'block', marginBottom: 4 }}>
              규칙 이름 *
            </label>
            <input
              type="text"
              value={formData.name}
              onChange={(e) => handleFieldChange('name', e.target.value)}
              placeholder="예: 자재비 합계 검증"
              style={{
                width: '100%',
                background: '#0d1117',
                border: '1px solid #30363d',
                borderRadius: 4,
                padding: '7px 10px',
                color: '#c9d1d9',
                fontSize: 12,
                outline: 'none',
                boxSizing: 'border-box',
              }}
              onFocus={(e) => e.target.style.borderColor = '#58a6ff'}
              onBlur={(e) => e.target.style.borderColor = '#30363d'}
            />
          </div>

          {/* 규칙 설명 입력 */}
          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 11, color: '#8b949e', display: 'block', marginBottom: 4 }}>
              설명
            </label>
            <input
              type="text"
              value={formData.description}
              onChange={(e) => handleFieldChange('description', e.target.value)}
              placeholder="규칙에 대한 간단한 설명"
              style={{
                width: '100%',
                background: '#0d1117',
                border: '1px solid #30363d',
                borderRadius: 4,
                padding: '7px 10px',
                color: '#c9d1d9',
                fontSize: 12,
                outline: 'none',
                boxSizing: 'border-box',
              }}
              onFocus={(e) => e.target.style.borderColor = '#58a6ff'}
              onBlur={(e) => e.target.style.borderColor = '#30363d'}
            />
          </div>

          {/* 템플릿 선택 */}
          <TemplateSelector
            selectedTemplate={selectedTemplate}
            onSelect={handleTemplateSelect}
          />

          {/* 블록 빌더 */}
          <BlockBuilder
            blocks={formData.blocks}
            onChange={(newBlocks) => handleFieldChange('blocks', newBlocks)}
          />

          {/* 심각도 선택 */}
          <div style={{ marginTop: 14, display: 'flex', alignItems: 'center', gap: 12 }}>
            <span style={{ fontSize: 11, color: '#8b949e' }}>심각도:</span>
            {SEVERITY_OPTIONS.map((sev) => (
              <span
                key={sev.id}
                className={`template-chip${formData.severity === sev.id ? ' selected' : ''}`}
                style={
                  formData.severity === sev.id
                    ? { borderColor: sev.color, color: sev.color, background: sev.borderColor }
                    : { borderColor: sev.color, color: sev.color }
                }
                onClick={() => handleFieldChange('severity', sev.id)}
              >
                {sev.label}
              </span>
            ))}
          </div>

          {/* 오류 메시지 */}
          {error && (
            <div style={{ marginTop: 10, fontSize: 11, color: '#f85149' }}>
              {error}
            </div>
          )}
        </div>

        {/* 모달 하단 버튼 */}
        <div className="modal-footer">
          {/* 삭제 버튼 (편집 시만 표시) */}
          {isEditing ? (
            <button
              className="btn-topbar"
              onClick={handleDelete}
              disabled={isSaving}
              style={{ color: '#f85149', borderColor: '#3d1a1a' }}
            >
              규칙 삭제
            </button>
          ) : (
            <div /> // 공간 유지
          )}

          <div className="btn-group">
            <button
              className="btn-topbar"
              onClick={handleTest}
              disabled={isSaving}
            >
              테스트 실행
            </button>
            <button
              className="btn-topbar"
              onClick={closeRuleEditor}
              disabled={isSaving}
            >
              취소
            </button>
            <button
              className="btn-topbar primary"
              onClick={handleSave}
              disabled={isSaving}
            >
              {isSaving ? '저장 중...' : '저장'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
