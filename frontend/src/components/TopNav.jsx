import useStore from '../stores/store.js';
import { useValidation } from '../hooks/useValidation.js';

// =============================================
// TopNav 컴포넌트 (TopBar 대체)
// 48px 블러 네비게이션 바
// 브랜드 로고 + 파일명 + 액션 버튼
// =============================================

export default function TopNav() {
  const isConnected = useStore((s) => s.isConnected);
  const activeWorkbook = useStore((s) => s.excelData.activeWorkbook);
  const { isValidating, runValidation, exportResult, isExporting } = useValidation();

  // 파일명 추출
  const fileName = activeWorkbook
    ? (typeof activeWorkbook === 'string'
        ? activeWorkbook.split(/[\\/]/).pop()
        : activeWorkbook.name || '파일명 없음')
    : '연결된 파일 없음';

  return (
    <div className="topnav">
      {/* 브랜드 */}
      <div className="brand">
        <span className="brand-icon">✓</span>
        Excel Validator
      </div>

      {/* 구분선 */}
      <div className="sep" />

      {/* 파일명 */}
      <div className="file">
        <span
          className="file-dot"
          style={{ background: isConnected ? 'var(--green)' : 'var(--s3)' }}
          title={isConnected ? 'Excel 연결됨' : 'Excel 연결 끊김'}
        />
        {fileName}
      </div>

      <div className="spacer" />

      {/* 액션 버튼 */}
      <div className="actions">
        <button
          className="pill-btn ghost"
          onClick={exportResult}
          disabled={isExporting}
          title="검증 결과를 JSON 파일로 저장"
        >
          {isExporting ? '저장 중...' : '내보내기'}
        </button>
        <button
          className="pill-btn accent"
          onClick={runValidation}
          disabled={isValidating || !activeWorkbook}
          title="모든 규칙으로 검증 재실행"
        >
          {isValidating ? '검증 중...' : '재검증'}
        </button>
      </div>
    </div>
  );
}
