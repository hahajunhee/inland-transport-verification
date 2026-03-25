@echo off
chcp 65001 > nul
echo.
echo ============================================
echo    내륙운송정산검증 시스템 시작 중...
echo ============================================
echo.
echo 브라우저가 자동으로 열립니다.
echo 서버를 종료하려면 이 창을 닫거나 Ctrl+C 를 누르세요.
echo.
start "" "http://127.0.0.1:8000"
python main.py
pause
