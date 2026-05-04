# 설비일보 Word 크롤러 설계

## 개요

DRM 보호된 Word 설비일보 파일을 사용자가 연 상태에서 win32com으로 크롤링하여 SQLite에 전처리 저장. CSV 내보내기를 통해 SQream에 이관. 최종 목표는 SQream에서 키워드 기반 검색.

## 기술 스택

| 구성 요소 | 기술 |
|----------|------|
| Word 접근 | win32com (pywin32) |
| 시스템 트레이 | pystray |
| 트레이 아이콘 | Pillow |
| UI (팝업/테이블) | tkinter |
| DB | SQLite → CSV → SQream 이관 |
| DB 파일 | `data/facility_daily.db` |

## 아키텍처

```
적재자가 Word 파일 열기
  → win32com 폴링 (4초) 자동 감지
  → 가장 큰 표 = 메인 표로 식별
  → 헤더(1행) 동적 읽기
  → 셀 내용을 2+ 개행으로 항목 분리
  → content_hash 기반 중복 체크
  → 모드에 따라 즉시 저장 or 팝업 확인 후 저장
  → SQLite 적재
  → CSV 내보내기 (날짜 범위 선택)
```

## 사용자 워크플로우

1. 적재 전담자가 `word_crawler.py` 실행 → 트레이 아이콘 상주
2. 확정된 설비일보 Word 파일을 순차적으로 열기
3. 자동 감지 → 파싱 → 모드에 따라 즉시 저장 or 팝업 확인
4. 중복 파일은 content_hash로 자동 필터링
5. 작업 완료 후 CSV 내보내기 → SQream COPY FROM

## DB 스키마

```sql
CREATE TABLE facility_daily (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    date             TEXT NOT NULL,
    source_file      TEXT NOT NULL,
    row_num          INTEGER NOT NULL,
    header1          TEXT,
    val1             TEXT,
    content_col_name TEXT,
    item_text        TEXT,
    raw_cell         TEXT,
    header4          TEXT,
    val4             TEXT,
    content_hash     TEXT,
    created_at       TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX idx_fd_date ON facility_daily(date);
CREATE INDEX idx_fd_source ON facility_daily(source_file);
CREATE INDEX idx_fd_hash ON facility_daily(content_hash);
```

### 컬럼 설명

| 컬럼 | 설명 |
|------|------|
| date | 문서 날짜 (표 위 텍스트 → 파일명 → 사용자 입력 순으로 추출) |
| source_file | 원본 Word 파일명 |
| row_num | 메인 표 내 행 번호 (2행부터, 1행은 헤더) |
| header1 | 첫 번째 열 헤더명 (동적) |
| val1 | 첫 번째 열 값 (예: Day, Night, Off 등) |
| content_col_name | 항목이 속한 열의 헤더명 (예: A동, B동 등) |
| item_text | 2+ 개행으로 분리된 개별 항목 |
| raw_cell | 해당 셀 전체 원문 (문맥 확인용) |
| header4 | 마지막 열 헤더명 (동적) |
| val4 | 마지막 열 값 (비고 등) |
| content_hash | 중복 방지용 해시 (SHA256 앞 16자) |

## 파싱 로직

### 표 식별
- 문서 내 모든 표를 순회
- **행 수가 가장 많은 표** = 메인 표로 식별
- 헤더 하드코딩 없음 (동적)

### 헤더 읽기
- 1행을 동적으로 읽어 열 이름 결정
- 4열 고정: col1(맥락) + col2, col3(내용) + col4(맥락)

### 셀 내용 분리
- 2회 이상 개행 (`\n\n+`)으로 항목 분리
- 각 항목 → 개별 레코드로 저장
- 태그 `[현상][원인][조치]` → 있으면 보너스 추출, 없어도 raw 저장
- 설비 키워드 태그 추출 없음

### 날짜 추출 (우선순위)
1. 표 위 텍스트에서 정규식 추출
2. 파일명에서 정규식 추출
3. 둘 다 실패 → 사용자 입력 팝업

### 저장 구조
- 비정규화 단일 테이블
- 한 셀에 항목 3개 → 3행 생성, 각 행에 raw_cell 동일 복사
- 같은 행의 col1, col4 값도 각 항목 레코드에 포함

## 중복 처리

- `content_hash` = SHA256(shift:building:raw_text 조합)[:16]
- `source_file` + `date` + `content_hash` 조합으로 판별
- 신규: 파싱 결과 표시/저장
- 동일: 토스트 알림 "이미 파싱된 파일입니다"
- 변경: 팝업에서 사용자 판단

## 저장 모드

트레이 우클릭 메뉴에서 전환:

| 모드 | 동작 |
|------|------|
| **확인 후 저장 (기본)** | 파싱 → 팝업 표시 → [저장] / [스킵] 선택 |
| **즉시 저장** | 파싱 → 바로 SQLite 저장 → 토스트 알림 |

## 파싱 실패 처리

| 상황 | 처리 |
|------|------|
| 표 없음 | "파싱 대상 아님" 토스트 알림 |
| 표 있으나 형식 불확실 | 최선 파싱 → 팝업에서 사용자 판단 (저장/스킵) |

## UI 구성

### 시스템 트레이
- pystray 아이콘 상주
- 우클릭 메뉴: 파싱 뷰어 열기 / 저장 모드 전환 / 종료
- 토스트 알림 3종: 성공(초록) / 중복(노랑) / 실패(빨강)

### 메인 팝업 (파싱 뷰어)
- **표시 조건**: ① 트레이 → "파싱 뷰어 열기" ② 확인 후 저장 모드에서 새 파싱 시 자동
- **상단**: 전체 이력/신규 건수 정보
- **좌측 패널**: 날짜 목록 (신규 = 초록 표시) + 전체 보기 버튼
- **중앙 테이블**: 날짜, 구분, 영역, 항목(item_text), 원문(raw_cell), 비고
- **하단 좌**: 선택 행 삭제
- **하단 우**: 스킵 / 저장
- **행 삭제**: 저장 전/후 모두 가능

### CSV 내보내기 다이얼로그
- 범위: 날짜 지정 / 전체 내보내기
- 시작일/종료일 선택
- 파일명 미리보기: `yyyy_mm_dd_yyyy_mm_dd.csv`
- content_hash 포함 → SQream 중복 방지

## SQream 이관

- SQLite → CSV 내보내기 → SQream `COPY FROM`
- content_hash로 중복 방지:
```sql
INSERT INTO facility_daily
SELECT * FROM staging
WHERE NOT EXISTS (
    SELECT 1 FROM facility_daily f
    WHERE f.content_hash = staging.content_hash
      AND f.source_file = staging.source_file
);
```

## 의존성

```
pystray>=0.19.5
pywin32>=306
Pillow>=10.0.0
```

## UI 목업

`mockup_word_crawler_ui.html` 참조
