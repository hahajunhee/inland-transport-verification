# 내륙운송정산검증 시스템

물류 회사의 내륙운송 정산 엑셀을 업로드하면 등록된 요율 기준과 자동 비교하여 과·부족 청구를 즉시 검출하는 로컬 웹 애플리케이션.

## 기술 스택

- **Backend**: Python 3.11+ / FastAPI 0.115 / Uvicorn
- **Database**: SQLite (WAL 모드, 단일 파일 `data/inland_transport.db`)
- **Template**: Jinja2 (서버 사이드 렌더링)
- **Frontend**: Vanilla JS + CSS (프레임워크 없음)
- **Excel 처리**: pandas + openpyxl
- **패키지 관리**: pip (`requirements.txt`)

## 폴더 구조

```
내륙운송정산/
├── main.py                  # FastAPI 앱 엔트리포인트 (uvicorn)
├── requirements.txt         # Python 의존성
├── install.bat / start.bat  # Windows 설치·실행 배치
├── app/
│   ├── data_store.py        # SQLite 데이터 저장소 (load/save/init_db)
│   ├── schemas.py           # Pydantic 스키마
│   ├── routers/             # FastAPI 라우터 (페이지 + API)
│   │   ├── pages.py         # HTML 페이지 라우터 (/, /verification, /mobis 등)
│   │   ├── rates.py         # 요율 CRUD API
│   │   ├── verification.py  # 2단계 정산검증 API
│   │   ├── mobis.py         # 3단계 모비스검증 API
│   │   ├── checklist.py     # 1단계 체크리스트 API
│   │   ├── trkv.py          # TRKV 구간요율 API
│   │   ├── storage_rates.py # 보관료/상하차료/셔틀료 요율 API
│   │   └── backup.py        # 요율 백업/복원 API
│   └── services/            # 비즈니스 로직 (라우터와 분리)
│       ├── verification_service.py  # 정산검증 핵심 로직
│       ├── mobis_service.py         # 모비스 검증 로직
│       ├── excel_service.py         # 엑셀 파싱·리포트 생성
│       ├── checklist_service.py     # 체크리스트 검증 로직
│       ├── rate_service.py          # 요율 관련 로직
│       ├── storage_rate_service.py  # 보관료 요율 로직
│       └── trkv_service.py          # TRKV 요율 로직
├── templates/               # Jinja2 HTML 템플릿
│   ├── base.html            # 공통 레이아웃 (navbar 포함)
│   ├── index.html           # 대시보드
│   ├── verification.html    # 2단계 정산검증 UI
│   ├── mobis.html           # 3단계 모비스검증 UI
│   ├── checklist.html       # 1단계 체크리스트 UI
│   └── rate_register.html   # 요율등록 통합 UI
├── static/
│   ├── css/style.css        # 전체 스타일시트
│   └── js/                  # 페이지별 JS (verification.js, trkv.js 등)
├── data/                    # 런타임 데이터 (git 제외)
│   └── inland_transport.db  # SQLite DB 파일
├── vba_modules/             # Excel VBA 모듈 (거래관리 매크로)
└── docs/                    # 문서
```

## 주요 명령어

```bash
# 의존성 설치
pip install -r requirements.txt

# 개발 서버 실행 (자동 리로드)
python main.py
# → http://127.0.0.1:8000

# Windows 배치 (설치 + 실행)
install.bat    # 최초 설치
start.bat      # 서버 시작 + 브라우저 자동 열기
```

## 환경변수

이 프로젝트는 환경변수를 사용하지 않음. 모든 설정은 코드 내 상수 또는 SQLite DB에 저장.

## 페이지 구조 & API 라우팅

| 페이지 | URL | API prefix | 설명 |
|--------|-----|------------|------|
| 대시보드 | `/` | - | 요약 통계 |
| 요율등록 | `/rate-register` | `/api/rates`, `/api/trkv`, `/api/storage-rates` | TRKV·보관료·상하차료·셔틀료 요율 등록 |
| 1단계 체크리스트 | `/checklist` | `/api/checklist` | 정산 엑셀 사전 점검 |
| 2단계 정산검증 | `/verification` | `/api/verification` | 요율 기준 자동 비교 검증 |
| 3단계 모비스검증 | `/mobis` | `/api/mobis` | GROVE↔MOBIS 금액 교차 검증 |
| 백업/복원 | - | `/api/backup`, `/api/restore` | 요율 데이터 JSON 내보내기/가져오기 |

## 아키텍처 규칙

- **비즈니스 로직은 services/ 에만**: 라우터는 요청 파싱 + 서비스 호출 + 응답만 담당
- **데이터 접근은 data_store 모듈 통해서만**: `data_store.load("파일명.json")`, `data_store.save("파일명.json", data)` — 파일명은 테이블명으로 매핑됨
- **프론트엔드에 빌드 도구 없음**: 순수 HTML/CSS/JS, 수정 즉시 반영
- **엑셀 파싱 패턴**: `pd.read_excel(BytesIO(bytes), header=None, dtype=str)` → 헤더를 코드에서 직접 탐색 (merged cell 대응)
- **SQLite WAL 모드**: 동시 읽기 허용, 쓰기는 `_write_lock` 으로 직렬화

## 모비스 검증 특이사항

- MOBIS 엑셀은 1:2행 병합 헤더 사용. 비용 열(내륙운임·보관료·상하차료·셔틀료·대기료) 탐색에 클러스터 기반 알고리즘 적용
- `_HEADER_ALIASES` 로 엑셀 오타 대응 (예: `상하자료` → `상하차료`)
- 2단계 정산검증 세션 참조 기능: `session_id` 로 기존 검증결과의 예상금액·구간정보를 키매칭

## 데이터 이전

요율 DB (`data/inland_transport.db`)는 git에 포함되지 않음:
1. **파일 복사**: `data/inland_transport.db` 를 그대로 복사
2. **JSON 백업**: 대시보드(`/`)에서 "요율 백업" → 새 PC에서 "요율 복원"

## 건드리지 말 것

- `data_store.py` 의 `_FILE_TO_TABLE` 매핑 — 기존 호환성 유지 필수
- `install.bat` / `start.bat` — Windows 운영 환경 필수 파일
- `vba_modules/` — Excel VBA 매크로 (별도 프로젝트, 참조용으로만 포함)
