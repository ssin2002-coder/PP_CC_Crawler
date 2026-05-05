import useStore from '../stores/store.js';
import { useValidation } from '../hooks/useValidation.js';

// =============================================
// 상단 바 컴포넌트
// 앱 타이틀, 파일명, 검증 결과 배지, 액션 버튼
// =============================================

export default function TopBar() {
  const isConnected = useStore((s) => s.isConnected);
  const activeWorkbook = useStore((s) => s.excelData.activeWorkbook);
  const { summary, isValidating, runValidation, exportResult, isExporting } = useValidation();

  // 파일명 추출 (객체 또는 문자열 대응)
  const fileName = activeWorkbook
    ? (typeof activeWorkbook === 'string' ? activeWorkbook.split(/[\\/]/).pop() : activeWorkbook.name || '파일명 없음')
    : '연결된 파일 없음';

  return (
    <div className="topbar">
      <div className="topbar-left">
        {/* 연결 상태 표시 점 */}
        <span
          className="connection-dot"
          title={isConnected ? 'Excel 연결됨' : 'Excel 연결 끊김'}
          style={{ background: isConnected ? '#3fb950' : '#6e7681' }}
        />
        <span className="app-title">Excel Validator</span>
        <span className="file-name" title={activeWorkbook || ''}>
          {fileName}
        </span>
      </div>

      <div className="topbar-right">
        {/* 오류 배지 */}
        {summary.errors > 0 && (
          <span className="status-badge badge-errors">
            {summary.errors} errors
          </span>
        )}
        {/* 경고 배지 */}
        {summary.warnings > 0 && (
          <span className="status-badge badge-warnings">
            {summary.warnings} warnings
          </span>
        )}
        {/* OK 배지 */}
        {summary.ok > 0 && (
          <span className="status-badge badge-ok">
            {summary.ok} ok
          </span>
        )}
        {/* 검증 전 상태 - 아무 배지도 없을 때 */}
        {summary.errors === 0 && summary.warnings === 0 && summary.ok === 0 && (
          <span className="status-badge" style={{ background: '#21262d', color: '#8b949e' }}>
            검증 대기중
          </span>
        )}

        {/* 결과 저장 버튼 */}
        <button
          className="btn-topbar"
          onClick={exportResult}
          disabled={isExporting}
          title="검증 결과를 JSON 파일로 저장"
        >
          {isExporting ? '저장 중...' : '결과 저장'}
        </button>

        {/* 검증 재실행 버튼 */}
        <button
          className="btn-topbar primary"
          onClick={runValidation}
          disabled={isValidating || !activeWorkbook}
          title="모든 규칙으로 검증 재실행"
        >
          {isValidating ? '검증 중...' : '검증 재실행'}
        </button>
      </div>
    </div>
  );
}
