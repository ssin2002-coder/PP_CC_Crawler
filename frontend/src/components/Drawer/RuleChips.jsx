import useStore from '../../stores/store.js';

// =============================================
// RuleChips 컴포넌트 (RuleSidebar 대체)
// 수평 줄바꿈 pill 칩 목록
// 클릭으로 규칙 on/off 토글
// =============================================

// 규칙 심각도에 따른 dot 클래스 결정
function getDotClass(rule) {
  if (!rule.enabled) return '';
  if (rule.severity === 'error') return 'de';
  if (rule.severity === 'warning' || rule.severity === 'warn') return 'dw';
  return 'do';
}

export default function RuleChips() {
  const rules = useStore((s) => s.rules);
  const toggleRule = useStore((s) => s.toggleRule);

  if (rules.length === 0) return null;

  return (
    <div className="rules-bar">
      {rules.map((rule) => {
        const dotClass = getDotClass(rule);

        return (
          <div
            key={rule.id}
            className={`rc${rule.enabled ? ' on' : ''}`}
            onClick={() => toggleRule(rule.id)}
            title={rule.description || rule.name}
          >
            <span className={`dot${dotClass ? ` ${dotClass}` : ''}`} />
            {rule.name}
          </div>
        );
      })}
    </div>
  );
}
