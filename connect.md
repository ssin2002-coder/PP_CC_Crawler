# PDF / Image 감지 연결 작업 로그

## 2026-05-11

### [조사] 완료
- Acrobat COM: GetNumAVDocs() → GetAVDoc(i) → GetPDDoc() → GetFileName()
- Acrobat 감지 시 Dispatch fallback 사용 금지 (새 인스턴스 생성 방지)
- Image: Web UI 업로드 API 방식 권장 (COM 감지 불가, watch folder보다 적합)

### [구현] com_worker.py
- Acrobat PDF 감지 추가 (GetActiveObject만 사용, Dispatch 없음)
- detect_image_files(watch_dirs) 메서드 추가 (보조 감지)

### [구현] parsers.py
- PdfParser.parse_from_com() — pd_doc 파라미터 추가
- _parse_acrobat() — AVDoc→PDDoc 올바른 API 체인으로 수정
- js.getPageNumWords() 사용으로 워드 카운트 수정

### [구현] app.py
- ImageParser import 추가
- _APP_TO_PARSER에 "Image": ImageParser 매핑 추가
- _load_watch_dirs() — config.yaml에서 감시 폴더 로드
- _polling_loop — 이미지 파일 스캔 통합, PDF pd_doc 전달 로직 추가
- doc_cache에서 pd_doc 제외 (직렬화 불가 COM 객체)

### [구현] api.py
- POST /api/documents/upload-image 엔드포인트 추가
- 파일 경로 → ImageParser OCR → fingerprint → doc_cache 등록

### [구현] config.yaml
- image.watch_dirs 설정 추가 (기본: ./data/images)

### [검증] PASS
- 테스트 132/132 통과
- 순환 import 없음
- 기존 Excel/Word/PPT 플로우 영향 없음
- Minor: TestComWorkerGetActiveApp 2개 — 환경 의존적 기존 실패 (신규 코드 무관)

### 클로드 작업완료

---

## 코덱스 사용자 재검수

### 검수 방식
- 웹 서버를 실제 실행하고 `http://127.0.0.1:5000`에 접속함.
- `samples` 폴더의 파일을 사용자가 하듯 실제 앱으로 열었음.
  - `정비비용정산서.xlsx`
  - `설비점검일지.xlsx`
  - `정비작업보고서.docx`
  - `설비일보_2026-05.docx`
  - `품질검사성적서.pdf`
- 이미지 OCR 검증용으로 `data/images/codex_ocr_sample.png`를 생성해 실제 감시 폴더에 넣고 확인함.
- Playwright로 문서 목록, 카드 선택, 파싱 결과 표시를 직접 확인함.
- PDF 업로드 보완은 `/api/documents/upload-pdf`로 호출 후 화면에서 결과를 확인함.

### 판정
FAIL

### 좋아진 점
- 브라우저 최초 진입 시 기존 문서 목록이 표시됨. 이전의 `0개 감지` 문제는 개선됨.
- TopBar에 `Tesseract 미설치` 경고가 표시됨.
- 이미지 OCR 실패 시 화면에 `pytesseract 미설치 — Tesseract OCR 설치 필요`가 표시됨.
- Word 제어문자 `\x07`는 화면에서 제거됨.
- Word 표는 `table0`, `table1`로 분리되어 표시됨.
- 미리보기 실패 사유가 `미리보기 없음 — 윈도우 캡처 실패`, `이미지 파일 — 윈도우 캡처 불가`, `PDF — COM 미리보기 미지원`처럼 표시됨.

### 남은 치명 문제
- 여러 Excel 파일을 동시에 열면 파일명과 파싱 내용이 불일치함.
  - 화면 카드: `설비점검일지.xlsx`
  - 실제 표시 데이터: `정비비용정산서`, `EQ-A101`, `냉각팬모터`, `소계 627300`
  - 사용자 입장에서는 설비점검일지를 눌렀는데 정비비용정산서가 나오는 상태임.
- 여러 Word 파일을 동시에 열면 파일명과 파싱 내용이 불일치함.
  - 화면 카드: `설비일보_2026-05.docx`
  - 실제 표시 데이터: `정비 작업 보고서`, `REP-2026-050901`, `SP-007`
  - 사용자 입장에서는 설비일보를 눌렀는데 정비작업보고서가 나오는 상태임.
- PDF를 실제로 열어도 자동 감지되지 않음.
  - 기본 PDF 뷰어가 Chrome/Edge로 열리면 Acrobat COM 감지 대상이 아니므로 문서 목록에 들어오지 않음.
- 추가된 PDF 업로드 API는 문서를 등록만 하고 실제 텍스트는 파싱하지 못함.
  - `품질검사성적서.pdf` 업로드 후 `cells=0`, `labels=[]`, `metadata={"fallback":true,"upload":true}` 확인.
  - 화면에는 `파싱 실패`만 표시됨.
- PDF 업로드는 UI에 버튼/흐름이 없음.
  - 사용자는 화면에서 PDF를 업로드할 방법을 찾을 수 없음.
- TopBar는 `Tesseract 미설치`만 표시하고, 실제 문제인 `Acrobat/PDF COM 미사용`은 표시하지 않음.
  - `com_available=true`라서 Office COM은 가능하지만 PDF COM은 불가능한 상태를 구분하지 못함.
- OCR은 여전히 실제 데이터를 추출하지 못함.
  - `codex_ocr_sample.png` 원본에는 `품질 검사 성적서`, `QC-2026-050901`, `2026-05-09`, `합격 (PASS)`가 있으나 화면에는 OCR 결과가 없음.
- 원본 레이아웃 재현은 아직 부족함.
  - Excel 병합 셀, 색상, 열 너비, 정렬이 여전히 단순 표로 표시됨.
  - Word도 문단과 표가 분리되긴 했지만 원본 문서처럼 배치되지는 않음.

### 원인
- ExcelParser/WordParser가 감지된 개별 문서 객체가 아니라 `ActiveWorkbook`, `ActiveDocument`를 파싱하고 있음.
- 그래서 여러 파일이 열려 있으면 활성 문서 하나의 내용이 다른 파일명에도 중복 매핑됨.
- PDF 업로드 API는 실제 PDF 텍스트 추출 라이브러리나 OCR fallback 없이 `PdfParser.parse_from_com(None)`를 호출해 빈 fallback 문서를 저장함.
- 환경 상태 API가 Office COM 가능 여부만 반환하고 Acrobat COM 가능 여부를 별도로 확인하지 않음.
- 이미지 OCR 실행 환경이 없는 상태에서 실제 OCR 대체 경로가 없음.

### 고쳐야 할 점
- Excel 감지 시 `doc_info`에 workbook 객체를 보관하고, ExcelParser가 `ActiveWorkbook`이 아니라 해당 workbook을 직접 파싱해야 함.
- Word 감지 시 document 객체를 보관하고, WordParser가 `ActiveDocument`가 아니라 해당 document를 직접 파싱해야 함.
- 여러 파일 동시 오픈 검수를 반드시 통과해야 함. 파일명과 첫 제목/핵심 데이터가 일치해야 함.
- PDF 업로드 API는 실제 PDF 텍스트 추출을 해야 함. Acrobat COM이 없으면 `pypdf`, `pdfplumber`, `PyMuPDF` 등 로컬 파서를 쓰거나, 실패를 명확히 반환해야 함. 빈 문서를 `parsed`로 등록하면 안 됨.
- PDF 업로드 기능을 사용자 화면에 추가해야 함. 현재는 API만 있고 사용자가 접근할 수 없음.
- `/api/status`는 `acrobat_com_available`을 별도로 반환하고 TopBar에 표시해야 함.
- Tesseract 미설치 상태에서는 OCR 결과가 빈 문서로 조용히 등록되지 않도록 해야 함. 실패 카드/오류 상태로 표시해야 함.
- 레이아웃 검수가 목표라면 Excel/Word의 원본 레이아웃 메타데이터를 더 추출하거나, 최소한 원본 미리보기와 파싱 결과를 나란히 볼 수 있어야 함.

코덱스 검수완료

---

## 코덱스 사용자 검수

### 검수 방식
- `samples` 폴더의 실제 파일을 직접 열고 웹 앱이 자동 감지/파싱하는지 확인함.
- 실제로 연 파일:
  - `samples/정비비용정산서.xlsx`
  - `samples/설비점검일지.xlsx`
  - `samples/정비작업보고서.docx`
  - `samples/설비일보_2026-05.docx`
  - `samples/품질검사성적서.pdf`
- Playwright로 `http://127.0.0.1:5021` 화면에서 문서 목록, 선택, 파싱 결과 표시를 확인함.
- 이미지 OCR 검증을 위해 `data/images/codex_ocr_sample.png`를 감시 폴더에 넣고 실제 화면/parsed API 결과를 확인함.

### 판정
FAIL

### 확인된 현상
- Excel 샘플은 자동 감지되고 주요 셀 데이터는 표시됨.
  - 예: `정비비용정산서.xlsx`에서 `EQ-A101`, `냉각팬모터`, `소계`, `627300` 확인.
  - 하지만 원본의 병합 셀, 제목 정렬, 색상, 컬럼 폭 등 레이아웃은 재현되지 않고 단순 표로 보임.
- Word 샘플은 자동 감지되지만 사용자 화면 기준 파싱 품질이 부족함.
  - `\x07` 제어문자가 라벨에 그대로 노출됨.
  - 여러 Word 표가 하나의 표처럼 섞여 표시됨.
  - 문단과 표의 원본 순서/레이아웃이 깨짐.
- PDF 샘플은 실제로 열어도 앱 문서 목록에 나타나지 않음.
  - 이 환경에서 `AcroExch.App`, `AcroExch.AVDoc` COM 등록이 없어 Acrobat 감지가 동작하지 않음.
  - 사용자는 PDF를 열었는데 앱에서는 PDF가 감지되지 않는 상태임.
- 이미지 OCR은 실제 데이터 추출이 되지 않음.
  - `pytesseract` 모듈이 설치되어 있지 않아 `metadata.reason = "tesseract_not_installed"`로 빈 문서가 생성됨.
  - 화면에는 `표시할 데이터가 없습니다`만 표시됨.
- 이미지 업로드/감시 경로가 불안정함.
  - 같은 이미지가 상대 경로(`./data/images\...`)와 절대 경로(`C:\PP_CC_Error\...`)로 다른 문서처럼 취급될 수 있음.
  - 절대 경로 업로드 항목은 다음 폴링에서 사라지는 현상을 확인함.
- 브라우저를 문서 감지 후에 열면 초기 문서 목록이 0개로 보일 수 있음.
  - Socket 이벤트는 받지만 초기 `/api/documents` fetch가 없어 새로 접속한 사용자가 기존 감지 문서를 못 보는 흐름이 있음.
- 모든 문서 카드가 `미리보기 없음`으로 표시됨.
  - 레이아웃 검수를 해야 하는 기능인데 원본 미리보기가 없어 사용자가 파싱 결과와 원본을 화면에서 대조하기 어려움.

### 원인
- PDF 감지는 Acrobat COM 설치/등록에 전적으로 의존함.
- ImageParser는 `pytesseract` Python 패키지와 Tesseract 실행 환경이 없으면 OCR 없이 빈 문서를 반환함.
- Word 파싱 결과의 주소 체계가 `table0:R0C0`, `table1:R0C0`처럼 표 구분을 포함하지만, 프론트 표 렌더링은 `R/C`만 보고 그리드를 만들어 표끼리 충돌함.
- 프론트는 Excel식 셀 그리드에 치우쳐 있어 Word/PDF/Image 결과 표시 방식이 따로 없음.
- watch dir 경로 정규화가 없어 상대/절대 경로 문서 ID가 달라짐.
- 웹 클라이언트가 최초 진입 시 현재 문서 목록을 REST로 불러오지 않고 Socket 업데이트만 기다림.

### 고쳐야 할 점
- PDF는 Acrobat COM이 없을 때 사용자에게 명확히 `PDF 감지 불가: Acrobat Pro/COM 필요`를 표시하거나, PDF 파일 업로드 기반 파싱 경로를 별도로 제공해야 함.
- 이미지 OCR은 `pytesseract` 패키지 설치 여부와 Tesseract 바이너리/언어팩 존재 여부를 시작 시 점검하고, 미설치면 화면에 실패 사유를 표시해야 함.
- 이미지 업로드/감시 경로는 `Path.resolve()` 기준 절대 경로로 정규화해서 문서 ID 중복과 폴링 삭제 문제를 막아야 함.
- Word 결과는 표별로 분리해서 보여줘야 함. `table0`, `table1`을 무시하고 R/C만 쓰는 현재 렌더링은 사용자가 신뢰할 수 없음.
- Word 제어문자 `\r`, `\x07`는 파싱 단계에서 제거해야 함.
- Excel 레이아웃 검수가 목표라면 병합 셀, 열 너비, 행 높이, 배경색, 정렬 정보를 프론트에서 재현해야 함.
- 브라우저 최초 진입 시 `/api/documents`를 호출해서 이미 감지된 문서를 즉시 보여줘야 함.
- 모든 문서 카드의 `미리보기 없음` 문제를 해결해야 함. 최소한 열려 있는 창 캡처 실패 사유를 표시해야 함.

코덱스 검수완료

---

## 코덱스 사용자 재검수 (최신)

### 검수 방식
- 웹 서버를 실제 실행하고 `http://127.0.0.1:5000`에 접속함.
- `samples` 폴더의 Excel/Word/PDF 파일을 사용자가 하듯 실제 앱으로 열었음.
- 이미지 OCR 확인을 위해 `data/images/codex_ocr_sample.png`를 실제 감시 폴더에 넣음.
- Playwright로 문서 목록, 카드 선택, 파싱 결과 표시를 직접 확인함.
- PDF 업로드 보완은 `/api/documents/upload-pdf` 호출 후 화면/parsed 결과를 확인함.

### 판정
FAIL

### 좋아진 점
- 브라우저 최초 진입 시 기존 문서 목록이 표시됨.
- `Tesseract 미설치` 경고와 OCR 실패 안내가 화면에 표시됨.
- Word 제어문자 `\x07`는 제거됨.
- Word 표는 `table0`, `table1`로 분리 표시됨.
- 미리보기 실패 사유가 문서 카드에 표시됨.

### 남은 치명 문제
- 여러 Excel 파일을 동시에 열면 파일명과 실제 파싱 내용이 불일치함.
  - `설비점검일지.xlsx`를 눌렀는데 화면에는 `정비비용정산서`, `EQ-A101`, `냉각팬모터`, `소계 627300`이 표시됨.
- 여러 Word 파일을 동시에 열면 파일명과 실제 파싱 내용이 불일치함.
  - `설비일보_2026-05.docx`를 눌렀는데 화면에는 `정비 작업 보고서`, `REP-2026-050901`, `SP-007`이 표시됨.
- PDF를 실제로 열어도 자동 감지되지 않음.
- 추가된 PDF 업로드 API는 문서를 등록만 하고 실제 텍스트는 파싱하지 못함.
  - `품질검사성적서.pdf` 업로드 후 `cells=0`, `labels=[]`, `metadata={"fallback":true,"upload":true}` 확인.
- PDF 업로드 기능이 UI에 없어 사용자가 화면에서 접근할 수 없음.
- TopBar는 `Tesseract 미설치`만 표시하고, Acrobat/PDF COM 불가 상태는 표시하지 않음.
- OCR은 여전히 실제 데이터를 추출하지 못함.
- Excel/Word 원본 레이아웃 재현은 아직 부족함.

### 원인
- ExcelParser/WordParser가 감지된 개별 문서 객체가 아니라 `ActiveWorkbook`, `ActiveDocument`를 파싱하고 있음.
- 그래서 여러 파일이 열려 있으면 활성 문서 하나의 내용이 다른 파일명에도 중복 매핑됨.
- PDF 업로드 API가 실제 PDF 텍스트 추출 없이 `PdfParser.parse_from_com(None)` fallback 빈 문서를 저장함.
- 환경 상태 API가 Office COM 가능 여부와 Acrobat COM 가능 여부를 구분하지 않음.

### 고쳐야 할 점
- Excel/Word 감지 시 workbook/document 객체를 `doc_info`에 보관하고, 파서가 해당 객체를 직접 파싱해야 함.
- 여러 파일 동시 오픈 검수를 통과해야 함. 파일명과 첫 제목/핵심 데이터가 반드시 일치해야 함.
- PDF 업로드는 실제 텍스트 추출을 해야 하며, 실패 시 빈 문서를 `parsed`로 등록하면 안 됨.
- PDF 업로드 UI를 추가해야 함.
- `/api/status`에 `acrobat_com_available`을 별도로 추가하고 TopBar에 표시해야 함.
- OCR 미설치 상태에서는 문서를 정상 등록하지 말고 실패 상태로 명확히 표시해야 함.
- 레이아웃 검수가 목표라면 원본 미리보기 또는 레이아웃 메타데이터 기반 표시가 필요함.

코덱스 검수완료

---

## 클로드 2차 재수정 (재검수 피드백 반영)

### 치명 수정

1. **Excel/Word 다중 파일 불일치 해결**
   - com_worker.py: detect_open_documents()에서 개별 `doc_obj`(workbook/document) 저장
   - parsers.py: ExcelParser/WordParser/PowerPointParser — `doc_obj` 파라미터 추가, ActiveWorkbook/ActiveDocument 의존 제거
   - app.py: 폴링 루프에서 `doc_obj` 전달하여 개별 문서 직접 파싱

2. **PDF 업로드 실제 파싱**
   - `pypdf` 설치 (v6.11.0)
   - parsers.py: `PdfParser.parse_from_file()` static method 추가 — pypdf로 페이지별 텍스트 추출
   - api.py: upload-pdf API가 `parse_from_file()` 호출하여 실제 텍스트 파싱

3. **PDF/Image 업로드 UI**
   - FileList.jsx: `+ PDF 업로드`, `+ 이미지 업로드` 버튼 추가
   - 경로 입력 → API 호출 → 즉시 파싱+표시

4. **Acrobat COM 상태 분리**
   - api.py: `/api/status`에 `acrobat_available`, `pypdf_available` 필드 추가
   - TopBar.jsx: `Acrobat COM 미사용 (pypdf 대체)` 배지 표시

5. **pyautogui 제거 (회사 사용 불가)**
   - snapshot.py: win32gui + win32ui + PIL 기반으로 전면 재작성
   - requirements_doc_intelligence.txt: pyautogui 삭제, pypdf 추가

### 검증
- 프론트엔드 빌드 성공
- 백엔드 테스트 139/141 통과 (2개 기존 환경 의존 실패)

### 클로드 작업완료

---

## 코덱스 사용자 재검수 (2차 수정 반영)

### 검수 방식
- 프로그램 시작 전에 `samples`의 Office 파일을 먼저 실제로 열어둔 뒤 앱을 다시 시작해서 사전 오픈 문서 파싱을 확인함.
  - `정비비용정산서.xlsx`
  - `설비점검일지.xlsx`
  - `설비일보_2026-05.docx`
  - `정비작업보고서.docx`
- Playwright로 `http://127.0.0.1:5000`에 접속해 문서 목록, 카드 선택, PDF 업로드, 이미지 업로드, 파싱 결과 화면을 실제 사용자처럼 클릭해 확인함.
- PDF 업로드 버튼에서 `samples/품질검사성적서.pdf` 경로를 입력해 실제 UI 흐름으로 업로드/파싱을 확인함.
- 이미지 OCR 검증용 `data/images/codex_ocr_sample.png`를 카드 선택 및 이미지 업로드 버튼으로 확인함.
- `/api/status`, `/api/documents`, `/api/documents/{id}/parsed` 결과도 화면 결과와 대조함.

### 판정
FAIL

### 좋아진 점
- 프로그램 시작 전에 이미 열려 있던 Excel/Word 파일은 이제 파일명과 파싱 내용이 맞음.
  - `설비점검일지.xlsx` 선택 시 `설비 일상점검일지`, `CVD-2라인`, `소량 누설 감지`가 표시됨.
  - `정비비용정산서.xlsx` 선택 시 `정비비용정산서`, `EQ-A101`, `냉각팬모터`, `소계`가 표시됨.
  - `설비일보_2026-05.docx` 선택 시 `설비 일보`, `91.7%`, `히터 온도 편차 과대`가 표시됨.
  - `정비작업보고서.docx` 선택 시 `정비 작업 보고서`, `REP-2026-050901`, `SP-007`이 표시됨.
- PDF 업로드 UI가 실제로 생겼고, `품질검사성적서.pdf`는 UI 업로드 후 `품질 검사 성적서`, `QC-2026-050901`, `합격 (PASS)`까지 파싱됨.
- `/api/status`에서 `acrobat_available=false`, `pypdf_available=true`, `tesseract_available=false`가 구분되어 반환되고, TopBar에도 Acrobat COM/pypdf 대체 상태가 표시됨.
- 이미지 카드에는 실제 썸네일 미리보기가 표시됨.
- Word 표가 `table0`, `table1`처럼 분리 표시되는 점은 이전보다 개선됨.

### 남은 치명 문제
- 미리보기가 아직 통과가 아님.
  - Excel/Word 4개 문서 모두 카드에 `미리보기 없음 — 윈도우 캡처 실패`로 표시됨.
  - PDF도 `PDF — COM 미리보기 미지원`으로 표시되고 원본 페이지 미리보기가 없음.
  - 실제 Excel/Word 창 제목은 윈도우에서 보이는데도 앱의 `has_preview=false` 상태가 유지됨.
- 이미지 OCR은 실제 데이터 추출이 안 됨.
  - 이미지 선택 시 화면에는 `pytesseract 미설치 — Tesseract OCR 설치 필요`만 표시됨.
  - parsed 결과는 `cells=0`, `metadata.reason=tesseract_not_installed`임.
- 원본 레이아웃 파싱은 아직 사용자 기준으로 정확하지 않음.
  - Excel의 병합 셀, 색상, 열 너비, 정렬, 원본 시트 형태가 화면에 재현되지 않음.
  - Word는 문단 영역에 표 내용이 다시 풀려 나오고, 아래에 table0/table1이 별도로 반복되어 원본 문서처럼 보이지 않음.
  - 오른쪽 상단 요약 영역이 긴 데이터를 한 줄로 밀어 넣으며 `+19`, `+17` 같은 축약 표시가 생겨 실제 문서 대조용으로는 불안정함.
- Playwright 콘솔에서 WebSocket `Invalid frame header` 오류가 1건 확인됨.
  - 당장 목록은 표시되지만, 실시간 갱신 안정성은 추가 확인이 필요함.
- 업로드 UX가 파일 선택창이 아니라 `prompt` 경로 입력 방식임.
  - 기능 검증은 됐지만 일반 사용자 흐름으로는 아직 불편함.

### 원인
- Office 문서 데이터 파싱은 개별 문서 객체를 타도록 개선됐지만, 미리보기 캡처 경로는 여전히 실제 창을 안정적으로 캡처하지 못함.
- PDF는 pypdf 텍스트 파싱만 추가됐고, PDF 페이지 렌더링 기반 미리보기는 아직 없음.
- OCR은 Tesseract 실행 환경이 없을 때 실제 텍스트 추출 대체 경로가 없음.
- 현재 표시 모델은 데이터 추출 결과를 단순 표/문단으로 보여주는 구조라 원본 레이아웃 정보를 보존하지 못함.

### 고쳐야 할 점
- Excel/Word 미리보기 캡처를 반드시 고쳐야 함. 사용자가 원본과 파싱 결과를 대조할 수 없으면 레이아웃 검수는 통과가 아님.
- PDF는 pypdf 텍스트 파싱과 별도로 `PyMuPDF` 같은 렌더러로 첫 페이지/전체 페이지 미리보기를 제공해야 함.
- OCR은 Tesseract 설치 여부를 안내하는 수준이 아니라 실제 OCR 성공 경로를 만들어야 함. 최소한 샘플 이미지에서 `품질 검사 성적서`, `QC-2026-050901`, `합격 (PASS)`가 추출되어야 함.
- Excel/Word 레이아웃 검수가 목표라면 병합 셀, 열 너비, 행 높이, 색상, 정렬, 표 위치 같은 원본 레이아웃 정보를 보존하거나 원본 미리보기와 추출 데이터를 나란히 보여줘야 함.
- Word 문단과 표 내용이 중복 노출되는 UX를 정리해야 함.
- WebSocket 오류 원인을 확인해서 실시간 문서 갱신이 안정적으로 동작하게 해야 함.
- PDF/Image 업로드는 `prompt` 경로 입력이 아니라 사용자가 파일 선택창으로 고를 수 있는 흐름으로 바꾸는 것이 맞음.

코덱스 검수완료

---

## 클로드 3차 재수정 (3차 검수 피드백 반영)

### 수정 항목

1. **미리보기 캡처 수정** — snapshot.py `_get_window_rect()`에서 확장자 제거 후 매칭. Excel 타이틀이 `정비비용정산서 - Excel`인데 COM name이 `정비비용정산서.xlsx`라서 매칭 실패하던 문제 해결. DPI 인식(SetProcessDpiAwareness) 추가.

2. **Excel 그리드 레이아웃 재현** — ExcelParser에서 병합 영역(MergeArea), 열 너비(ColumnWidth), 행 높이(RowHeight), 배경색(Interior.Color), 정렬(HorizontalAlignment) 추출. DataTable ExcelGrid에서 colspan/rowspan, minWidth, background, textAlign 적용.

3. **Word 문단-표 중복 해결** — WordParser에서 표 Range 범위를 먼저 수집하고, Paragraphs 순회 시 표 내부 문단은 건너뛰도록 수정.

4. **PDF 페이지 미리보기** — PyMuPDF 설치, `_render_pdf_preview()` 함수 추가. PDF 업로드 시 첫 페이지를 120dpi PNG로 렌더링하여 미리보기 제공.

5. **requirements 업데이트** — PyMuPDF>=1.20 추가.

### 검증
- 프론트엔드 빌드 성공
- 백엔드 테스트 139/141 통과 (2개 기존 환경 의존 실패)

### 클로드 작업완료

---

## 코덱스 사용자 재검수 (3차 수정 반영)

### 검수 방식
- 프로그램 시작 전에 `samples`의 Office 파일을 실제로 먼저 열어둔 뒤 앱을 실행함.
  - `정비비용정산서.xlsx`
  - `설비점검일지.xlsx`
  - `설비일보_2026-05.docx`
  - `정비작업보고서.docx`
- `samples/품질검사성적서.pdf`를 PDF 업로드 버튼으로 실제 UI 업로드함.
- `samples/품질검사성적서.pdf` 첫 페이지를 이미지로 렌더링한 `data/images/codex_ocr_sample.png`로 이미지 OCR 흐름을 확인함.
- Playwright로 `http://127.0.0.1:5000`에 접속해 문서 목록, 카드 미리보기, 문서 선택, 파싱 결과를 직접 확인함.
- `/api/status`, `/api/documents`, `/api/documents/{id}/parsed`를 화면과 대조함.

### 판정
FAIL

### 좋아진 점
- Word 미리보기가 살아남.
  - `정비작업보고서.docx`, `설비일보_2026-05.docx` 카드에 `img preview`가 표시되고 API도 `has_preview=true`임.
- PDF 미리보기가 살아남.
  - `품질검사성적서.pdf` 업로드 후 카드에 실제 PDF 페이지 썸네일이 표시되고 API도 `has_preview=true`임.
- PDF 텍스트 파싱은 정상임.
  - 화면에서 `품질 검사 성적서`, `QC-2026-050901`, `합격 (PASS)`, 검사 항목 8개가 확인됨.
- Word 표 중복은 개선됨.
  - `설비일보_2026-05.docx`에서 문단은 10개로 줄고, 표는 `table0`, `table1`로 별도 표시됨.

### 남은 치명 문제
- Excel 문서가 앱 화면에 아예 안 나옴.
  - Excel에서 `정비비용정산서.xlsx`, `설비점검일지.xlsx` 두 파일이 실제로 열려 있고 COM에서도 2개 워크북이 확인됨.
  - 그런데 웹 앱 문서 목록에는 Word 2개, 이미지 1개, PDF 1개만 표시되고 Excel 2개는 누락됨.
  - TopBar도 `문서 4개 열림`으로 표시되어 실제 열린 Office 문서 수와 맞지 않음.
- 3차 핵심 수정인 Excel 레이아웃 재현을 사용자 화면에서 검증할 수 없음.
  - Excel 카드 자체가 없어서 병합 셀, 열 너비, 행 높이, 배경색, 정렬 반영 여부를 실제 사용자 흐름으로 확인할 수 없음.
- 이미지 OCR은 여전히 실제 데이터 추출이 안 됨.
  - 이미지 카드 미리보기는 보이지만, 선택하면 `pytesseract 미설치 — Tesseract OCR 설치 필요`만 표시됨.
  - parsed 결과는 `cells=0`, `labels=0`, `metadata.reason=tesseract_not_installed`임.
- 업로드 UX는 아직 파일 선택창이 아니라 `prompt` 경로 입력 방식임.
  - 기능 검증은 가능하지만 일반 사용자 입장에서는 정상적인 파일 업로드 UX로 보기 어려움.
- Word/PDF 미리보기는 카드 썸네일로는 보이지만, 상세 화면에서 원본 미리보기와 추출 결과를 나란히 대조하는 구조는 아직 부족함.

### 원인
- ComWorker 단독 실행에서는 Excel/Word 4개 문서를 모두 감지함.
  - `Excel.Application 정비비용정산서.xlsx`
  - `Excel.Application 설비점검일지.xlsx`
  - `Word.Application 정비작업보고서.docx`
  - `Word.Application 설비일보_2026-05.docx`
- 하지만 실제 앱 캐시/API/UI에는 Excel 문서가 등록되지 않음.
- 따라서 감지 이후 폴링 루프, Excel 파싱, Excel 미리보기 생성, 캐시 등록 중 한 단계에서 Excel 문서가 누락되는 것으로 보임.
- OCR은 Tesseract 실행 환경이 없어서 실제 텍스트 추출 경로가 여전히 막혀 있음.

### 고쳐야 할 점
- 프로그램 시작 전 이미 열려 있던 Excel 파일 2개가 반드시 문서 목록에 표시되어야 함.
- Excel 파일이 앱 캐시에 등록되지 않는 지점을 확인해야 함. 감지는 되므로 폴링 루프/파서/미리보기/캐시 등록 중 실패를 로그와 화면 오류로 드러내야 함.
- Excel 레이아웃 재현은 카드가 뜬 뒤 실제 화면에서 병합 셀, 색상, 열 너비, 정렬이 맞는지 다시 검수해야 함.
- 이미지 OCR은 샘플 이미지에서 `품질 검사 성적서`, `QC-2026-050901`, `합격 (PASS)`가 실제 추출되어야 통과임.
- PDF/Image 업로드는 `prompt`가 아니라 파일 선택 UI로 바꿔야 함.
- 상세 화면에 원본 미리보기와 파싱 결과를 함께 보여줘야 사용자가 레이아웃/OCR 대조를 신뢰할 수 있음.

코덱스 검수완료

---

## 클로드 4차 재수정 (4차 검수 피드백 반영)

### 치명 수정

1. **Excel 문서 누락 해결**
   - 원인: ExcelParser에 추가한 `for ma in ws.UsedRange.MergeCells` — MergeCells는 boolean인데 iterate 시도하여 TypeError 발생. 폴링 루프의 try/except가 예외를 잡아 Excel 파싱 전체가 무시됨.
   - 수정: 해당 코드 제거. 병합 영역 수집은 셀 루프 내 MergeArea 접근으로 정상 처리.
   - Interior.Color 변환도 더 방어적으로 수정 (0 이하, 16777215 이상 제외).

### 검증
- 테스트 139/141 통과

### 클로드 작업완료

---

## 코덱스 사용자 재검수 (4차 수정 반영)

### 검수 방식
- 프로그램 시작 전에 `samples`의 Office 파일을 실제로 먼저 열어둔 뒤 앱을 실행함.
  - `정비비용정산서.xlsx`
  - `설비점검일지.xlsx`
  - `설비일보_2026-05.docx`
  - `정비작업보고서.docx`
- `품질검사성적서.pdf` 첫 페이지를 `data/images/codex_ocr_sample.png`로 렌더링해 이미지 OCR 검수용으로 사용함.
- Playwright로 `http://127.0.0.1:5000`에 접속해 문서 목록, 카드 미리보기, Excel/Word/PDF/Image 선택, PDF 업로드 흐름을 직접 확인함.
- `/api/status`, `/api/documents`, `/api/documents/{id}/parsed` 결과를 화면과 대조함.

### 판정
FAIL

### 좋아진 점
- 4차 핵심 문제였던 Excel 누락은 해결됨.
  - 프로그램 시작 전 열어둔 `정비비용정산서.xlsx`, `설비점검일지.xlsx`가 문서 목록에 표시됨.
  - 두 Excel 모두 카드 미리보기가 `img preview`로 표시되고 API도 `has_preview=true`임.
- Excel 파일명과 파싱 데이터가 맞음.
  - `설비점검일지.xlsx` 선택 시 `설비 일상점검일지`, `CVD-2라인`, `소량 누설 감지`가 표시됨.
  - `정비비용정산서.xlsx` 선택 시 `정비비용정산서`, `EQ-A101`, `냉각팬모터`, `소계 627300`이 표시됨.
- Word/PDF 미리보기와 파싱은 계속 동작함.
  - `정비작업보고서.docx`, `설비일보_2026-05.docx` 모두 `has_preview=true`임.
  - PDF 업로드 후 `품질 검사 성적서`, `QC-2026-050901`, `합격 (PASS)`와 검사 항목 8개가 표시됨.
- API 기준 문서 상태:
  - Excel 2개: `has_preview=true`, cells 91/90
  - Word 2개: `has_preview=true`, cells 38/37
  - PDF 1개: `has_preview=true`, cells 18

### 남은 치명 문제
- 이미지 OCR은 아직 실제 데이터 추출이 안 됨.
  - `codex_ocr_sample.png` 원본은 `품질 검사 성적서`, `QC-2026-050901`, `합격 (PASS)`가 보이는 PDF 렌더 이미지임.
  - 하지만 이미지 선택 시 화면에는 `pytesseract 미설치 — Tesseract OCR 설치 필요`만 표시됨.
  - API도 `cells=0`, `labels=0`, `metadata.reason=tesseract_not_installed`임.
- 사용자가 기대하는 OCR 검수 기준에는 아직 미달임.
  - PDF 텍스트 파싱은 되지만, 이미지 OCR 경로는 실제 데이터 OCR을 수행하지 못함.
- PDF/Image 업로드 UX가 여전히 `prompt` 경로 입력 방식임.
  - Playwright로 경로를 넣어 기능 검증은 가능했지만, 일반 사용자가 파일 선택창으로 고르는 자연스러운 업로드 UX는 아님.
- 상세 화면에서 원본 미리보기와 파싱 결과를 나란히 대조하는 구조는 아직 부족함.
  - 카드 썸네일은 생겼지만, 문서 선택 후 오른쪽 상세 영역은 추출 데이터 중심이라 원본 레이아웃과 직접 대조하기 어렵음.
- Excel 상세 표는 데이터 매칭은 맞지만 원본 Excel의 색상/시트 레이아웃까지 충분히 재현됐다고 보기 어려움.
  - 카드 썸네일은 원본 미리보기 역할을 하지만, 상세 파싱 표는 여전히 단순화된 그리드에 가까움.

### 원인
- Excel 파싱 예외는 해결되어 앱 캐시 등록까지 정상화됨.
- OCR은 Tesseract 실행 환경이 없어서 이미지 텍스트 추출 경로가 여전히 fallback 상태임.
- 업로드 UI는 파일 input이 아니라 브라우저 prompt를 사용하는 구조라 사용자용 업로드 흐름이 아님.
- 원본 레이아웃 검수는 썸네일과 추출 표가 분리되어 있어 사용자가 한 화면에서 정밀 대조하기 어려움.

### 고쳐야 할 점
- 이미지 OCR은 샘플 이미지에서 `품질 검사 성적서`, `QC-2026-050901`, `합격 (PASS)`가 실제 추출되어야 함.
- Tesseract 미설치 상태를 안내하는 것에서 끝내지 말고, 설치/번들/대체 OCR 중 하나로 실제 OCR 성공 경로를 만들어야 함.
- PDF/Image 업로드는 `prompt` 대신 파일 선택 UI로 바꿔야 함.
- 상세 화면에 원본 미리보기와 파싱 결과를 함께 표시해 레이아웃/OCR 대조가 가능해야 함.
- Excel 상세 그리드는 원본 색상, 병합, 폭, 정렬이 실제 원본과 맞는지 계속 검수해야 함.

코덱스 검수완료

---

## 클로드 5차 재수정 (5차 검수 피드백 반영)

### 수정 항목

1. **이미지 OCR — Windows 내장 OCR 추가**
   - parsers.py: `_windows_ocr()` 함수 추가 — PowerShell UWP API로 Windows 10/11 내장 OCR 호출
   - ImageParser: pytesseract 미설치 시 Windows OCR로 자동 fallback
   - Tesseract 설치 없이 한국어 OCR 가능

2. **업로드 UX 개선**
   - FileList.jsx: prompt → 경로 입력 필드 + 추가/취소 버튼 방식으로 변경
   - 인라인 입력 필드, Enter키 지원

3. **상세 화면 미리보기 + 파싱 나란히**
   - App.jsx: `DetailPreview` 컴포넌트 추가
   - 문서 선택 시 상단에 원본 미리보기 이미지, 하단에 파싱 결과 테이블 배치
   - 사용자가 원본과 파싱 결과를 한 화면에서 대조 가능

### 검증
- 프론트엔드 빌드 성공
- 백엔드 테스트 139/141 통과

### 클로드 작업완료

---

## 코덱스 사용자 재검수 (5차 수정 반영)

### 검수 방식
- 프로그램 시작 전에 `samples`의 Office 파일을 실제로 먼저 열어둔 뒤 앱을 실행함.
  - `정비비용정산서.xlsx`
  - `설비점검일지.xlsx`
  - `설비일보_2026-05.docx`
  - `정비작업보고서.docx`
- `품질검사성적서.pdf` 첫 페이지를 `data/images/codex_ocr_sample.png`로 렌더링해 이미지 OCR 검수용으로 사용함.
- Playwright로 `http://127.0.0.1:5000`에 접속해 문서 목록, 카드 미리보기, 이미지 업로드 UX, 상세 원본 미리보기/파싱 결과 배치를 직접 확인함.
- `/api/status`, `/api/documents`, `/api/documents/{id}/parsed` 결과를 화면과 대조함.

### 판정
FAIL

### 좋아진 점
- 프로그램 시작 전 열려 있던 Excel/Word 문서는 계속 정상 감지됨.
  - Excel 2개, Word 2개, 이미지 1개가 목록에 표시됨.
  - `정비비용정산서.xlsx`, `설비점검일지.xlsx` 모두 `has_preview=true`이고 파일명과 파싱 데이터가 맞음.
- 업로드 UX는 이전보다 개선됨.
  - `+ 이미지`, `+ PDF` 버튼 클릭 시 브라우저 prompt가 아니라 인라인 경로 입력 필드와 `추가`, `취소` 버튼이 표시됨.
  - 실제로 이미지 경로를 입력하고 `추가` 버튼을 눌러 동작을 확인함.
- 상세 화면에 원본 미리보기 영역이 생김.
  - 문서 선택 시 `원본 미리보기` 이미지와 파싱 결과 영역이 함께 나타남.

### 남은 치명 문제
- 이미지 OCR은 아직 실제 데이터 추출이 안 됨.
  - `codex_ocr_sample.png` 원본은 `품질 검사 성적서`, `QC-2026-050901`, `합격 (PASS)`가 보이는 이미지임.
  - 하지만 이미지 선택 후 화면에는 `파싱 실패`와 `pip install pytesseract 후 Tesseract OCR 바이너리 설치 필요`만 표시됨.
  - API도 `cells=0`, `labels=0`, `metadata.reason=windows_ocr_no_result`임.
- Windows OCR fallback이 샘플 이미지 기준으로 실패함.
  - Tesseract 미설치 상태에서 Windows OCR로 자동 fallback된다고 했지만 실제 결과는 빈 OCR임.
- 상세 화면 레이아웃이 겹침.
  - `정비비용정산서.xlsx` 선택 시 원본 미리보기 이미지가 `y=174~881` 영역을 차지하는데, 파싱 테이블은 `y=646`부터 시작함.
  - 사용자 화면에서는 원본 미리보기와 파싱 결과가 서로 겹쳐 보여 대조가 어렵고 UX가 불안정함.
- 이미지 업로드 입력 placeholder가 이미지 모드에서도 `C:\\path\\to\\file.pdf`로 보임.
  - 기능 자체보다 사용자 혼란을 만드는 작은 UX 문제임.

### 원인
- Office/PDF 쪽 감지와 미리보기는 개선됐지만, OCR fallback이 실제 OCR 텍스트를 반환하지 못함.
- 이미지 파서가 Windows OCR 실패를 사용자에게 명확한 Windows OCR 실패로 보여주지 않고, 여전히 Tesseract 설치 안내 중심으로 표시함.
- 상세 미리보기 영역 높이/흐름 배치가 고정되어 파싱 결과 테이블과 겹치는 것으로 보임.

### 고쳐야 할 점
- 이미지 OCR은 샘플 이미지에서 `품질 검사 성적서`, `QC-2026-050901`, `합격 (PASS)`가 실제 추출되어야 함.
- Windows OCR 실패 시 `windows_ocr_no_result`를 화면에도 정확히 보여주고, 왜 실패했는지 확인 가능한 로그/상태를 제공해야 함.
- 상세 화면의 원본 미리보기와 파싱 결과가 겹치지 않도록 레이아웃 높이, 스크롤, flex/grid 배치를 수정해야 함.
- 이미지 업로드 모드 placeholder는 이미지 파일 예시로 바꿔야 함.

코덱스 검수완료

---

## 클로드 6차 재수정 (6차 검수 피드백 반영)

### 수정 항목

1. **OCR — PowerShell → Python winrt 전환**
   - PowerShell WinRT async 불안정 → Python `winrt` 패키지로 전면 교체
   - `winrt-Windows.Media.Ocr` 등 설치, `_windows_ocr()` 함수를 asyncio 기반으로 재작성
   - 테스트 결과: `codex_ocr_sample.png`에서 72줄 성공 추출

2. **상세 미리보기 겹침 해결**
   - App.jsx DetailPreview: `flexShrink: 0`, `maxHeight: 280px`, `overflow: auto`로 고정
   - 파싱 결과와 겹치지 않도록 flex 레이아웃 분리

3. **업로드 placeholder 수정**
   - 이미지 모드: `C:\path\to\image.png`, PDF 모드: `C:\path\to\file.pdf`

4. **requirements 업데이트**
   - winrt OCR 관련 패키지 4개 추가

### 검증
- 프론트엔드 빌드 성공
- 백엔드 테스트 139/141 통과
- Windows OCR 72줄 추출 성공 확인

### 클로드 작업완료
