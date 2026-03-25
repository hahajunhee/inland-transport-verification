# 내륙운송정산검증 시스템

운송 요율을 등록하고, 정산 엑셀 파일을 업로드하여 요율 기준 금액과의 차이를 즉시 확인할 수 있는 로컬 웹 애플리케이션입니다.

## 요구사항

- Python 3.11 이상
- Windows (배치 파일 기준)

## 설치 및 실행

### 처음 실행 시
```
install.bat
```
필요한 Python 패키지를 자동으로 설치합니다.

### 서버 시작
```
start.bat
```
서버가 시작되고 브라우저가 자동으로 열립니다. (`http://127.0.0.1:8000`)

### 수동 실행
```bash
pip install -r requirements.txt
python main.py
```

## 주요 기능

- **TRKV 요율 설정** (`/trkv`): 포트명 매핑, 컨테이너 티어, 구간별 요율 관리
- **일반 요율 관리** (`/rates`): 보관료, 상하차료, 셔틀비용 요율 등록
- **정산 검증** (`/verification`): 엑셀 파일 업로드 → 요율 기준 예상금액 계산 → 차이 확인 및 내보내기

## 데이터 이전 방법

요율 데이터(`data/transport.db`)는 깃에 포함되지 않습니다.

다른 PC로 이전할 때는 두 가지 방법 중 하나를 사용하세요:

1. **폴더 통째 복붙**: `data/transport.db` 파일이 포함되므로 요율이 그대로 유지됩니다.
2. **JSON 백업/복원**: 대시보드(`/`)에서 "요율 백업" 다운로드 → 다른 PC의 대시보드에서 "요율 복원" 업로드

## 기술 스택

- Backend: FastAPI + SQLAlchemy (SQLite)
- Frontend: Vanilla JS + Jinja2 Templates
