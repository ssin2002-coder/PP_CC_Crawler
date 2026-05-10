# Block 1: 문서 자동 감지 + 미리보기 + 핑거프린트 — 설계 문서

## 개요

Doc Intelligence의 첫 번째 레고 블록. 사용자가 열어놓은 문서(Excel/Word/PPT/PDF/Image)를 자동 감지하고, 웹 브라우저 UI에서 미리보기 + 파싱 데이터 + 템플릿 매칭 결과를 표시한다.

## 범위

- COM 기반 문서 자동 감지 (3초 폴링)
- 감지된 문서 스냅샷 캡처 (pyautogui 윈도우 캡처)
- 파싱 데이터 HTML 테이블 렌더링
- TF-IDF 핑거프린트 자동 매칭 + 학습 UI
- Flask + React 웹 UI (Apple Dark 테마)

## 범위 밖

- 파이프라인 캔버스 (블록 2)
- 드래그 영역 선택 (블록 3)
- 영역 비교 조건 연결 (블록 4)

## 파일 구조

```
doc_intelligence/
  __init__.py
  com_worker.py        ← 유지 (COM 감지/연결)
  parsers.py           ← 유지 (5종 파서)
  fingerprint.py       ← 유지 (TF-IDF 학습/매칭)
  engine.py            ← 유지 (데이터 모델: ParsedDocument, CellData, Fingerprint)
  storage.py           ← 유지 (SQLite CRUD)
  config.yaml          ← 유지

  web/                 ← 신규
    app.py             ← Flask + SocketIO 서버 + COM 폴링 스레드
    api.py             ← REST 엔드포인트
    snapshot.py        ← 윈도우 캡처 (pyautogui)
    frontend/          ← React (Vite)
      src/
        App.jsx
        components/
          TopBar.jsx
          FileList.jsx
          FileCard.jsx
          DocPreview.jsx
          FingerInfo.jsx
          DataTable.jsx
        stores/
          store.js     ← Zustand
        hooks/
          useSocket.js ← socket.io-client
      index.html
      vite.config.js
      package.json

run.py                 ← 진입점 (웹서버 시작 + 브라우저 오픈)
```

### 삭제 대상 (기존 doc_intelligence/)

- `main.py` — tkinter UI (웹으로 대체)
- `ui_components.py` — tkinter 위젯
- `anomaly.py` — Isolation Forest (블록 1 불필요)
- `graph.py` — NetworkX 시각화 (블록 1 불필요)
- `region_linker.py` — 드래그 영역 (블록 3에서 웹으로 재구현)
- `validator.py` — 검증 룰 (블록 4에서 재구현)
- `extractor.py` — 엔티티 추출 (블록 1 불필요. FingerInfo 필드 목록은 fingerprint.generate()의 labels로 충분)

## API 설계

### REST 엔드포인트

| Method | Path | 역할 |
|--------|------|------|
| GET | `/api/documents` | 감지된 열린 문서 목록 + 매칭 상태 |
| GET | `/api/documents/<id>/preview` | 스냅샷 이미지 (PNG base64) |
| GET | `/api/documents/<id>/parsed` | 파싱된 셀 데이터 (JSON) |
| POST | `/api/templates/learn` | 학습 요청 `{doc_id, template_name}` |
| POST | `/api/templates/confirm` | 후보 확인 `{doc_id, template_id}` |
| GET | `/api/templates` | 저장된 템플릿 목록 |

### WebSocket 이벤트 (SocketIO)

| 방향 | 이벤트 | 데이터 |
|------|--------|--------|
| Server → Client | `documents_updated` | 문서 목록 변경 시 (감지 결과 전체) |
| Server → Client | `parse_complete` | 파싱 완료 알림 `{doc_id, status}` |

## 문서 ID 정의

`doc_id`는 `file_path`의 MD5 해시(32자 hex)로 생성한다. 이는 `fingerprint.py`의 `_doc_id()` 방식과 동일하다.

```python
doc_id = hashlib.md5(file_path.encode("utf-8")).hexdigest()
```

## 서버 메모리 캐시

`web/app.py`는 다음 인메모리 캐시를 관리한다:

```python
_doc_cache: dict[str, dict] = {}
# key: doc_id
# value: {
#     "info": {...},              # detect_open_documents() 결과
#     "parsed": ParsedDocument,   # 파싱 결과 (learn() 호출 시 재사용)
#     "fingerprint": dict,        # fingerprint.generate() 결과
#     "match": dict,              # fingerprint.match() 결과
#     "snapshot_b64": str,        # PNG base64
#     "confirmed": bool,         # 후보 확인 여부 (재분류 방지)
# }
```

## COM 세션 관리

폴링 스레드에서는 `com_worker.com_session()` 컨텍스트 매니저를 사용한다:

```python
def _polling_loop(com_worker, interval=3):
    with com_worker.com_session():  # CoInitialize/CoUninitialize 보장
        while running:
            docs = com_worker.detect_open_documents()
            # ... 변경 감지 처리
            time.sleep(interval)
```

## 데이터 흐름

```
[COM 폴링 스레드] ─3초 간격─→ com_session() 내부에서 detect_open_documents()
       │ 변경 감지 시
       ▼
[Flask 서버]
  1. doc_info["app"] (ProgID) → _APP_TO_PARSER 매핑으로 파서 선택
     ┌─────────────────────────────────────────┐
     │ "Excel.Application"      → ExcelParser  │
     │ "Word.Application"       → WordParser   │
     │ "PowerPoint.Application" → PPTParser    │
     │ "AcroExch.App"           → PdfParser    │
     └─────────────────────────────────────────┘
  2. 선택된 파서.parse_from_com(com_app) → ParsedDocument
  3. fingerprint.generate() + fingerprint.match() 실행
  4. snapshot.capture_window() 실행
  5. _doc_cache에 저장 (ParsedDocument 포함)
  6. SocketIO emit('documents_updated')
       │
       ▼
[React UI]
  documents_updated 수신 → Zustand 업데이트 → FileList 리렌더
```

### PDF/Image 감지 한계

- PDF(`AcroExch.App`)는 `detect_open_documents()`에서 감지 시도하나, Adobe Reader 버전에 따라 COM 미등록 가능
- Image는 COM 감지 불가 → 블록 1 범위에서는 Excel/Word/PPT만 지원
- PDF/Image 지원은 향후 파일 드롭 또는 경로 직접 입력 방식으로 확장 예정

## UI 레이아웃

### Row 1: TopBar (고정)
- 좌: "Doc Intelligence" + 버전
- 우: 상태 도트(●) + "COM 연결됨 | 문서 N개 열림" + [+ 영역 연결] + [설정]

### Row 2: MainLayout (1:2 비율)

**좌측 패널 (33%) — FileList**
- PanelHeader: "열린 문서" + 카운트
- FileCard 반복:
  - 아이콘 + 파일명
  - 스냅샷 미리보기 이미지 (pyautogui 캡처)
  - 상태 뱃지: 매칭됨(초록) / 후보(주황) / 새 문서(파랑)

**우측 패널 (67%)**
- 상단 FingerInfo:
  - 매칭 결과: "✓ 정산서 양식 (92%)" 또는 "? 후보: 정산서 양식 (73%)"
  - 필드 목록: 태그 형태
  - 학습 버튼 (미매칭 시) / "이 양식인가요?" 확인 UI (후보 시)
- 하단 DataTable:
  - 시트 탭 (Excel일 경우)
  - 파싱된 셀 데이터를 HTML 테이블로 렌더링

## 학습 플로우

### 미매칭 (score < 0.60)
1. 뱃지: "새 문서" (파랑)
2. FingerInfo에 [학습] 버튼 표시
3. 클릭 → 모달: 템플릿 이름 입력
4. POST /api/templates/learn `{doc_id, template_name}`
   → `_doc_cache[doc_id]["parsed"]`에서 `ParsedDocument` 조회
   → `fingerprint.learn(doc, template_name)` 호출
   → `_doc_cache[doc_id]["match"]` 갱신 (auto=True)
5. 뱃지 → "매칭됨" (초록)으로 전환

### 후보 (0.60 ≤ score < 0.85)
1. 뱃지: "후보" (주황)
2. FingerInfo에 "이 양식인가요? [양식명]" + [예] [아니오] 표시
3. [예] → POST /api/templates/confirm → increment_match_count + _doc_cache에 confirmed=True 설정 → 뱃지 "매칭됨"
   - confirmed=True인 문서는 다음 폴링 때 재분류하지 않음
4. [아니오] → [학습] 버튼 표시 (미매칭 플로우와 동일)
   - 새 템플릿으로 학습. 기존 유사 템플릿은 유지 (삭제/병합 안 함)

### 자동 매칭 (score ≥ 0.85)
1. 뱃지: "매칭됨" (초록)
2. FingerInfo에 "✓ [양식명] ([score]%)" 표시

## 디자인 토큰 (Apple Dark)

| 역할 | 값 |
|------|-----|
| 배경 (메인) | `#000000` |
| 배경 (패널) | `#1d1d1f` |
| 배경 (카드) | `#1d1d1f` |
| 보더 | `#333336` |
| 텍스트 (메인) | `#f5f5f7` |
| 텍스트 (보조) | `#86868b` |
| 액센트 (블루) | `#0071e3` |
| 액센트 (밝은 블루) | `#2997ff` |
| 성공 (그린) | `#30d158` |
| 경고 (오렌지) | `#ff9f0a` |
| 모서리 반경 (카드) | `12px` |
| 모서리 반경 (버튼) | `980px` (pill) |
| 폰트 | Inter, system-ui |

## 기술 스택

| 레이어 | 기술 |
|--------|------|
| 서버 | Flask + flask-socketio (async_mode=threading) |
| COM | pywin32 (com_worker.py) |
| 파싱 | parsers.py + fingerprint.py |
| 스냅샷 | pyautogui |
| DB | SQLite (storage.py) |
| 프론트 | React 19 + Vite + Zustand + socket.io-client |
| 디자인 | Apple Dark 토큰 |

## 기존 코드 재활용

- `com_worker.py` — `detect_open_documents()`, `get_active_app()`, `com_session()` 그대로 사용
- `parsers.py` — `ExcelParser`, `WordParser` 등 `parse_from_com()` 그대로 사용
- `fingerprint.py` — `generate()`, `learn()`, `match()`, `process()` 그대로 사용
- `engine.py` — `ParsedDocument`, `CellData`, `Fingerprint` 데이터 모델 그대로 사용
- `storage.py` — SQLite 템플릿 CRUD 그대로 사용
