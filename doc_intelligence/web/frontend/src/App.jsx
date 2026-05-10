import TopBar from './components/TopBar';
import FileList from './components/FileList';
import FingerInfo from './components/FingerInfo';
import DataTable from './components/DataTable';
import { useStore } from './stores/store';
export default function App() {
  const selectedDocId = useStore((s) => s.selectedDocId);
  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <TopBar />
      <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
        <div style={{ width: '33.3%', borderRight: '1px solid var(--border)', overflow: 'auto' }}>
          <FileList />
        </div>
        <div style={{ width: '66.7%', display: 'flex', flexDirection: 'column', overflow: 'auto' }}>
          {selectedDocId ? (
            <>
              <FingerInfo />
              <DataTable />
            </>
          ) : (
            <div style={{
              flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center',
              color: 'var(--text-sub)', fontSize: '14px'
            }}>
              좌측에서 문서를 선택하세요
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
