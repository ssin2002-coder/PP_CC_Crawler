import { useEffect } from 'react';
import TopNav from './components/TopNav.jsx';
import HeroStrip from './components/HeroStrip.jsx';
import Drawer from './components/Drawer/Drawer.jsx';
import ExcelArea from './components/ExcelGrid/ExcelArea.jsx';
import { useWebSocket } from './hooks/useWebSocket.js';
import { useExcelData } from './hooks/useExcelData.js';
import useStore from './stores/store.js';

// =============================================
// 루트 App 컴포넌트
// 새 레이아웃: TopNav + HeroStrip + (Drawer + GridArea)
// BottomPanel 제거 → 이슈는 Drawer로 이동
// RuleEditorModal은 추후 활성화 예정
// =============================================

function AppInner() {
  // WebSocket 연결 초기화 (마운트 시 자동 연결)
  const { isConnected } = useWebSocket();

  // Excel 데이터 초기 로드
  useExcelData();

  return (
    <div className="app">
      {/* 상단 네비게이션 바 (48px) */}
      <TopNav />

      {/* 통계 띠 (110px) */}
      <HeroStrip />

      {/* 메인 영역: 서랍(330px) + Excel 그리드(flex:1) */}
      <div className="main">
        {/* 좌측 서랍 패널 */}
        <Drawer />

        {/* 우측 Excel 그리드 영역 */}
        <ExcelArea />
      </div>
    </div>
  );
}

export default function App() {
  return <AppInner />;
}
