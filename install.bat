@echo off
chcp 65001 > nul
echo.
echo ============================================
echo    내륙운송정산검증 - 패키지 설치
echo ============================================
echo.
echo [1/2] Python 버전 확인 중...
python --version
if errorlevel 1 (
  echo.
  echo [오류] Python이 설치되지 않았습니다.
  echo https://www.python.org/downloads/ 에서 Python 3.11 이상을 설치하세요.
  pause
  exit /b 1
)
echo.
echo [2/2] 필요 패키지 설치 중...
pip install -r requirements.txt
echo.
echo ============================================
echo    설치 완료!
echo    start.bat 을 실행하여 서버를 시작하세요.
echo ============================================
echo.
pause
