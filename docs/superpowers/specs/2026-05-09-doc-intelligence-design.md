# Doc Intelligence - 정비비용정산 문서 지능 분석 시스템

## 1. 개요

### 1.1 목적
정비비용정산 관련 문서(Excel, Word, PowerPoint, PDF, 이미지)를 자동으로 파싱하고, 문서 간 연관관계를 분석하여 "원래 어떻게 정의되어야 하는데 무엇이 문제인지"를 자동으로 탐지하는 시스템.

### 1.2 핵심 문제
- 같은 "정비비용정산" 맥락이지만, 양식/구조가 업체별, 시기별, 문서종류별로 모두 다름
- 하나의 정산 건을 검증하려면 여러 파일(A, B, C...)을 교차 참조해야 함
- 각 파일에서 봐야 할 위치와 내용이 다름

### 1.3 핵심 혁신
- **템플릿 핑거프린팅**: 문서를 열 때마다 구조적 지문을 생성하여 같은 양식을 자동 판별
- **드래그 영역 연결**: 화면에 띄운 문서 위에 직접 드래그하여 비교 영역을 지정
- **단일 룰 + 프리셋 조합**: 개별 검증 룰을 레고 블록처럼 조합하여 정산 유형별 프리셋 생성
- **자동 학습**: 문서를 열수록 시스템이 점점 똑똑해지는 구조

### 1.4 제약 조건
- **폐쇄망 환경**: 인터넷 접근 불가
- **LLM 미사용**: 클라우드/로컬 LLM 모두 사용하지 않음. 사전학습 트랜스포머 모델(BERT 등)도 제외. 순수 Python 라이브러리 + 전통 ML(scikit-learn)만 사용
- **보안 제약**: 파일 직접 접근 불가 — 사용자가 열어놓은 파일을 COM(win32com)으로 파싱
- **DRM 보호**: 문서가 DRM으로 보호되어 있어 네이티브 앱의 COM 인터페이스를 통해서만 접근 가능
- **Adobe Acrobat Pro 필수**: PDF COM 파싱을 위해 Adobe Acrobat Pro가 설치되어 있어야 함. Acrobat Reader만 있는 환경에서는 PDF 파서 대신 화면 캡처 + OCR 폴백 사용

### 1.5 기존 시스템과의 관계

본 프로젝트(Doc Intelligence)는 기존 Excel Validator(Flask + React)를 **대체**하는 것이 아니라, 별도의 독립 시스템으로 구축한다.

- **기존 Excel Validator**: 단일 Excel 파일 내 셀 검증 (범위, 중복, 합계 등)
- **Doc Intelligence**: 다중 문서 간 교차 검증 + 양식 학습 + 연관관계 분석

기존 코드 재활용:
- `word_crawler.py`의 Word COM 파싱 로직 -> `parsers.py`의 WordParser에 흡수
- `excel_crawler.py`의 Excel COM 파싱 로직 -> `parsers.py`의 ExcelParser에 흡수
- `backend/excel_com_worker.py`의 COM 별도 프로세스 패턴 -> `engine.py`에서 COM 안정성 패턴 재활용
- `backend/validators/`의 검증 로직은 참조하되, 단일 룰 + 프리셋 구조로 재설계

## 2. 아키텍처

### 2.1 레고 블록 구조

코어 엔진 위에 독립 모듈을 하나씩 추가하는 플러그인 방식.
각 블록은 독립적으로 개발/테스트 가능하며, 다른 블록 없이도 코어 엔진에 연결하면 동작함.

```
[코어 엔진] ← 필수. 이것만으로도 동작
    |
    +-- [블록 1] 문서 파서 (Excel/Word/PPT/PDF/이미지 각각 독립)
    +-- [블록 2] 구조 핑거프린터
    +-- [블록 3] 엔티티 추출기
    +-- [블록 4] 템플릿 매칭 엔진
    +-- [블록 5] 교차 검증기 (단일 룰 + 프리셋)
    +-- [블록 6] 드래그 영역 연결기
    +-- [블록 7] 이상 탐지기
    +-- [블록 8] 관계 그래프
    +-- [블록 9] 대시보드 UI
```

### 2.2 파일 구조

```
doc_intelligence/
+-- main.py              # 진입점 + tkinter 메인 UI
+-- engine.py            # 파이프라인 오케스트레이터 + 플러그인 레지스트리
+-- parsers.py           # 모든 문서 파서 (COM 기반)
+-- fingerprint.py       # 구조 핑거프린트 생성 + 템플릿 매칭
+-- extractor.py         # 엔티티 추출 (regex + 형태소 분석)
+-- validator.py         # 교차 검증 (단일 룰 + 프리셋 조합)
+-- region_linker.py     # 드래그 영역 연결 (투명 오버레이)
+-- storage.py           # SQLite 메타데이터 + 템플릿 + 룰 저장
+-- ui_components.py     # tkinter 공통 위젯 (뷰어, 편집기, 트리뷰)
+-- templates.db         # 학습된 템플릿/룰 DB (자동 생성)
```

총 9개 파일, DB 1개, 외부 서비스 0개.

### 2.3 플러그인 인터페이스

각 블록이 코어 엔진에 연결되려면 다음 프로토콜을 구현해야 한다:

```python
class PluginProtocol:
    """모든 블록이 구현하는 최소 인터페이스"""
    name: str                    # 플러그인 이름
    enabled: bool = True         # 활성/비활성 토글

    def initialize(self, engine) -> None:
        """엔진에 등록될 때 호출"""
        ...

    def process(self, doc: ParsedDocument, context: dict) -> dict:
        """파이프라인에서 호출. 결과를 context에 추가하여 다음 블록에 전달"""
        ...
```

블록 등록: `engine.register(MyPlugin())` — 엔진이 등록된 블록 순서대로 실행.
블록 비활성화: `engine.disable("fingerprint")` — 해당 블록을 건너뜀.
블록이 None이거나 disabled면 해당 단계를 자동 스킵.

### 2.4 기술 스택

| 영역 | 라이브러리 | 용도 |
|------|-----------|------|
| COM 연동 | win32com (pywin32) | Excel/Word/PPT/PDF 앱 연결 |
| 한국어 형태소 | kiwipiepy | 형태소 분석, 신조어 자동 추출 |
| 키워드 추출 | YAKE | 외부 모델 불필요, 통계 기반 |
| 문서 유사도 | TF-IDF + cosine_similarity (scikit-learn) | 핑거프린트 벡터 비교 |
| 클러스터링 | DBSCAN 또는 AgglomerativeClustering (scikit-learn) | 유사 양식 자동 그룹핑 |
| 이상 탐지 | Isolation Forest (scikit-learn) | 통계적 이상 패턴 |
| 그래프 | NetworkX | 문서 간 연관관계 모델링 |
| OCR | Tesseract (pytesseract) | 이미지 텍스트 추출 |
| 화면 캡처 | pyautogui | 드래그 영역 스크린샷 |
| 윈도우 관리 | win32gui | 창 판별, 좌표 변환 |
| DB | SQLite (표준 라이브러리) | 템플릿/룰/문서 메타데이터 |
| UI | tkinter (표준 라이브러리) | 데스크톱 UI |

## 3. 상세 설계

### 3.1 공통 데이터 모델

```python
@dataclass
class ParsedDocument:
    """모든 파서가 반환하는 통일된 구조"""
    file_path: str
    file_type: str              # "excel" | "word" | "ppt" | "pdf" | "image"
    raw_text: str               # 전체 텍스트
    structure: dict             # 문서 구조 정보 (시트, 섹션, 슬라이드 등)
    cells: list[CellData]       # 셀/필드 단위 데이터 (위치 + 값)
    metadata: dict              # 파일명, 생성일, 수정일 등

@dataclass
class CellData:
    """문서 내 개별 데이터 단위"""
    address: str                # "Sheet1!C5", "p3:para2", "slide1:textbox3" 등
    value: str
    data_type: str              # "text" | "number" | "date" | "formula"
    neighbors: dict             # 인접 셀 정보 (위/아래/좌/우)

@dataclass
class Entity:
    """추출된 엔티티"""
    type: str                   # "금액" | "날짜" | "업체명" | "설비코드" 등
    value: str
    location: str               # 문서 내 위치
    confidence: float           # 0.0 ~ 1.0

@dataclass
class Fingerprint:
    """문서 구조 지문"""
    doc_id: str
    feature_vector: list[float]
    label_positions: dict       # 키 레이블의 위치 맵
    merge_pattern: str          # 셀 병합 패턴 해시
```

### 3.2 문서 파서 (parsers.py)

모든 파서는 COM을 통해 사용자가 열어놓은 파일에 접근.

| 파일 유형 | COM 객체 | 추출 항목 |
|----------|---------|----------|
| Excel | Excel.Application | 셀 데이터, 시트 구조, 수식, 병합 패턴 |
| Word | Word.Application | 문단, 테이블, 인라인 이미지 |
| PowerPoint | PowerPoint.Application | 슬라이드, 텍스트박스, 표, 도형 텍스트 |
| PDF | AcroExch.App (Adobe Acrobat Pro 필수). Acrobat Pro 미설치 시 화면 캡처 + OCR로 폴백 | 텍스트, 테이블, 레이아웃 |
| 이미지 | 화면 캡처 + Tesseract OCR. 한국어 인쇄체 기준 정확도 80~90%. 저해상도/필기체는 정확도 저하 -> confidence 임계값(0.6) 미달 시 사용자 보정 요청 | OCR 텍스트, 블록 좌표 |

```python
class BaseParser:
    def parse_from_com(self, com_app) -> ParsedDocument: ...

class ExcelParser(BaseParser):
    """Excel.Application COM에서 활성 워크북 파싱"""
    # 기존 excel_crawler.py 로직 확장
    # 시트별 셀 데이터 + 병합 정보 + 수식 추출

class WordParser(BaseParser):
    """Word.Application COM에서 활성 문서 파싱"""
    # 기존 word_crawler.py 로직 확장
    # 문단 + 테이블 + 인라인 이미지 추출

class PowerPointParser(BaseParser):
    """PowerPoint.Application COM에서 활성 프레젠테이션 파싱"""
    # 슬라이드별 텍스트박스, 표, 도형 텍스트 추출

class PdfParser(BaseParser):
    """AcroExch.App COM에서 활성 PDF 파싱"""
    # 텍스트, 테이블, 레이아웃 추출

class ImageParser(BaseParser):
    """화면 캡처 후 Tesseract OCR"""
    # OCR 텍스트, 블록 좌표 추출
```

### 3.3 핑거프린트 + 템플릿 매칭 (fingerprint.py)

문서의 "뼈대"를 벡터화하여 같은 양식인지 자동 판별.

추출 특성:
- 고정 텍스트 레이블 해시 ("견적서", "합계", "업체명" 등)
- 레이아웃 좌표 벡터 (레이블이 어디에 있는가)
- 구조 패턴 (셀 병합, 테이블 크기, 시트/슬라이드/페이지 수)

매칭 임계값:
- >= 0.85: 자동 매칭 -> 바로 엔티티 추출
- 0.60 ~ 0.84: 후보 제시 -> 사용자 확인
- < 0.60: 새 양식 -> 학습 모드 진입

### 3.4 엔티티 추출 (extractor.py)

두 가지 모드:
1. **템플릿 모드**: 매칭된 템플릿의 정해진 위치에서 정확히 추출
2. **자동 분석 모드**: 새 양식일 때 regex + 형태소 분석으로 추정

정규식 패턴:
- 금액: `[\d,]+\s*원`
- 날짜: `\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}`
- 설비코드: `[A-Z]{2,3}-\d{3,5}`
- 사업자번호: `\d{3}-\d{2}-\d{5}`
- 전화번호: `\d{2,3}-\d{3,4}-\d{4}`

인접 셀 기반 추정: "합계" 옆 셀 -> 금액일 가능성 높음

### 3.5 교차 검증: 단일 룰 + 프리셋 (validator.py)

#### 단일 룰
개별 검증 규칙. 재사용 가능한 최소 단위.

룰 타입:
- **값_일치**: A의 영역 값 = B의 영역 값
- **순서_확인**: 날짜1 < 날짜2 < 날짜3
- **범위_확인**: 값이 지정 범위 내에 있는지
- **수식_확인**: 수량 x 단가 = 소계
- **존재_확인**: 필수 문서/필드가 있는지
- **포함_확인**: A의 값이 B에 포함되는지

```python
class SingleRule:
    id: int
    name: str                    # "금액일치", "날짜순서" 등
    rule_type: str               # "값_일치" | "순서_확인" | ...
    regions: list[LinkedRegion]  # 드래그로 지정된 영역들
    params: dict                 # 룰별 추가 파라미터
```

#### 프리셋
단일 룰들의 조합. 정산 유형별로 하나씩.

```python
class Preset:
    id: int
    name: str           # "정산A (배관정비)"
    category: str       # "배관", "전기", "설비" 등
    rule_ids: list[int] # [1, 2, 3, 5] — 포함된 룰 ID들
```

같은 룰이 여러 프리셋에서 재사용 가능.

실행 흐름:
1. 프리셋 선택, 또는 자동 감지: 열린 문서들의 template_id 조합을 presets 테이블의 연관 템플릿 목록과 매칭하여 후보 프리셋 제안. 복수 후보 시 사용자에게 선택 요청
2. 해당 프리셋의 룰들을 순서대로 실행
3. 결과: 통과 / 실패 / 경고 리스트 -> validation_results 테이블에 저장

### 3.6 드래그 영역 연결 (region_linker.py)

사용자가 화면에 띄운 문서 위에 직접 드래그하여 비교 영역을 지정하는 UX.

구현:
1. tkinter 투명 오버레이 창을 화면 전체에 띄움
   - WS_EX_TRANSPARENT 윈도우 스타일로 아래 앱의 포커스 유지
   - win32api.SetWindowPos()로 topmost 관리
2. 사용자가 문서 위에서 드래그 -> 영역 캡처
3. win32gui.WindowFromPoint()로 어떤 앱 창 위인지 판별
4. DPI 스케일링 처리:
   - ctypes.windll.shcore.GetDpiForMonitor()로 현재 모니터 DPI 획득
   - 스크린 좌표를 앱별 논리 좌표로 변환 (scale factor 적용)
   - 멀티모니터 환경: 각 모니터별 DPI 개별 처리
5. 앱별 좌표 -> 문서 위치 변환:
   - Excel: ActiveWindow.RangeFromPoint(x, y) -> 셀 주소
   - Word: ActiveWindow.Document.Range에서 좌표 기반 Selection 생성 (Word COM에 RangeFromPoint 없음 -> 화면 캡처 + 영역 좌표 저장 방식으로 대체)
   - PowerPoint: ActiveWindow.PointsToScreenPixels 역변환 -> 슬라이드 내 Shape 좌표
   - PDF: AcroExch.AVPageView.GetPageNum + DevPtToPagePt -> 페이지 내 좌표
   - 이미지: 스크린 좌표 그대로 저장 (OCR 텍스트 블록과 매핑)
6. 선택된 영역들을 연결하여 단일 룰 생성
7. 룰 타입 선택 팝업 -> 저장

사용자 경험:
1. [영역 연결] 버튼 클릭
2. 화면이 살짝 어두워짐 (오버레이)
3. 문서 A 위에서 드래그 -> 파란 테두리
4. 문서 B 위에서 드래그 -> 파란 테두리
5. 문서 C 위에서 드래그 -> 파란 테두리
6. 팝업: "관계 유형?" -> [값이 같아야 함] [합계] [날짜 순서] [사용자 정의]
7. 룰 이름 입력 -> 저장
8. 프리셋에 추가

### 3.7 저장소 (storage.py)

SQLite 단일 파일 DB (templates.db).

테이블 4개:

```sql
-- 학습된 문서 양식
CREATE TABLE templates (
    id INTEGER PRIMARY KEY,
    name TEXT,
    file_type TEXT,
    fingerprint_vector BLOB,
    label_positions TEXT,      -- JSON
    field_mappings TEXT,       -- JSON: 어떤 위치에 어떤 엔티티가 있는지
    created_at TIMESTAMP,
    match_count INTEGER DEFAULT 0
);

-- 단일 검증 룰
CREATE TABLE rules (
    id INTEGER PRIMARY KEY,
    name TEXT,
    rule_type TEXT,
    regions TEXT,               -- JSON: 연결된 영역 정보
    params TEXT,                -- JSON: 룰별 파라미터
    created_at TIMESTAMP
);

-- 프리셋 (룰 조합)
CREATE TABLE presets (
    id INTEGER PRIMARY KEY,
    name TEXT,
    category TEXT,
    rule_ids TEXT,               -- JSON: [1, 2, 3, 5]
    template_ids TEXT,           -- JSON: 연관 템플릿 ID 목록 (자동 감지용)
    created_at TIMESTAMP
);

-- 파싱된 문서 기록
CREATE TABLE documents (
    id INTEGER PRIMARY KEY,
    file_path TEXT,
    file_type TEXT,
    template_id INTEGER,
    entities TEXT,               -- JSON: 추출된 엔티티들
    parsed_at TIMESTAMP,
    FOREIGN KEY (template_id) REFERENCES templates(id)
);

-- 검증 결과 이력
CREATE TABLE validation_results (
    id INTEGER PRIMARY KEY,
    preset_id INTEGER,
    rule_id INTEGER,
    document_ids TEXT,           -- JSON: 관련 문서 ID 목록
    status TEXT,                 -- "통과" | "실패" | "경고"
    detail TEXT,                 -- JSON: 상세 내용 (기대값, 실제값, 위치 등)
    executed_at TIMESTAMP,
    FOREIGN KEY (preset_id) REFERENCES presets(id),
    FOREIGN KEY (rule_id) REFERENCES rules(id)
);
```

### 3.8 학습 모드 UX

새 양식 문서를 열었을 때의 흐름:

1. 파싱 -> 핑거프린트 생성 -> 매칭 실패
2. 자동 분석 결과를 UI에 표시
   - "이 셀이 '금액'인 것 같습니다" (confidence: 0.7)
   - "이 영역이 '업체명'인 것 같습니다" (confidence: 0.8)
3. 사용자가 틀린 부분만 수정
   - 드롭다운으로 필드 유형 변경
   - 영역 드래그로 범위 수정
4. "이 양식 저장" -> 새 템플릿으로 등록
5. 다음에 같은 양식 문서 -> 자동 매칭

학습 결과의 적용 범위:
- 사용자 보정 결과는 **해당 템플릿에만 적용**됨 (필드 위치 매핑 저장)
- 전역 자동 분석 로직(regex 패턴, 인접 셀 추정)은 변경하지 않음
- 단, 사용자가 반복적으로 추가하는 새 엔티티 패턴은 config.yaml의 커스텀 패턴 목록에 수동 추가 가능

### 3.9 COM 안정성 전략

COM 연동은 본질적으로 불안정하므로 다음 전략을 적용:

- **별도 프로세스 격리**: 기존 excel_com_worker.py 패턴을 재활용. COM 호출을 별도 프로세스에서 실행하여 메인 UI 스레드 보호
- **STA(Single-Threaded Apartment)**: COM 호출은 반드시 STA 스레드에서 실행 (pythoncom.CoInitialize)
- **재시도 정책**: COM 호출 실패 시 최대 3회 재시도 (1초 간격)
- **타임아웃**: COM 호출 10초 타임아웃. 초과 시 프로세스 종료 후 사용자에게 알림
- **문서 닫힘 감지**: 파싱 중 사용자가 문서를 닫으면 COM 에러 캐치 후 "문서가 닫혔습니다" 알림
- **RPC 실패 복구**: pywintypes.com_error 캐치 -> COM 앱 재연결 시도

### 3.10 메인 UI (main.py + ui_components.py)

tkinter 기반 데스크톱 앱. 기존 크롤러 UX 패턴 유지.

tkinter 선택 근거:
- 투명 오버레이(드래그 영역 연결)가 웹 기반에서는 OS 레벨 창 제어 불가
- 기존 word_crawler/excel_crawler가 tkinter 기반 -> 일관성 유지
- 단일 실행파일 배포 용이 (PyInstaller)
- 폐쇄망에서 별도 웹 서버 불필요

주요 화면:
- **메인 창**: 활성 문서 모니터링, 파싱 상태, 최근 결과
- **학습 모드**: 문서 미리보기 + 자동 분석 결과 편집
- **룰 관리**: 단일 룰 목록 + 프리셋 관리 (체크박스 조합)
- **영역 연결**: 투명 오버레이 드래그 모드
- **검증 결과**: 통과/실패/경고 리스트 + 관련 문서 링크

## 4. 구현 순서

단계적 레고 블록 추가 방식:

| 단계 | 블록 | 산출물 |
|------|------|--------|
| MVP | 코어 엔진 + Excel 파서 + 핑거프린터 | "이 Excel은 이전에 본 양식 A와 같다" 판별 |
| +1 | 엔티티 추출기 | 금액, 날짜, 업체명 자동 추출 |
| +2 | 템플릿 매칭 + 학습 모드 UI | 새 양식 자동 분석 + 사용자 보정 |
| +3 | 드래그 영역 연결기 | 화면에서 드래그로 비교 영역 지정 |
| +4 | 단일 룰 + 프리셋 (교차 검증기) | 룰 조합으로 정산 검증 |
| +5 | Word 파서 추가 | Word 문서도 동일 파이프라인 |
| +6 | PowerPoint 파서 추가 | PPT도 동일 파이프라인 |
| +7 | PDF 파서 (Acrobat COM) | PDF도 동일 파이프라인 |
| +8 | 이미지 파서 + OCR | 사진/도면 OCR 처리 |
| +9 | 이상 탐지기 (Isolation Forest) | 통계적 이상 패턴 자동 감지 |
| +10 | 관계 그래프 (NetworkX + pyvis) | 문서 간 연결 시각화 |

## 5. 폐쇄망 배포

### 5.1 사전 준비 (인터넷 환경)
- Python 3.10+ 설치 파일
- pip wheelhouse: 모든 의존성 패키지의 Windows x64 wheel 파일
  - kiwipiepy는 C++ 확장이므로 반드시 플랫폼 맞는 wheel 준비
- Tesseract 바이너리 (v5.x) + 한국어 언어팩 (kor.traineddata)
- Adobe Acrobat Pro 설치 확인 (PDF 파서용)

### 5.2 폐쇄망 설치
```
pip install --no-index --find-links=./wheelhouse -r requirements.txt
```

### 5.3 필수 의존성
- pywin32 (win32com, win32gui, pythoncom)
- kiwipiepy (Windows x64 wheel 사전 빌드 필수)
- scikit-learn (>= 1.3)
- networkx
- pytesseract
- pyautogui
- yake
- pyvis (관계 그래프 시각화, +10 단계에서 추가)
