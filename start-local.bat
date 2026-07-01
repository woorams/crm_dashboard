@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PORT=13000
set DATA_DIR=%~dp0local-data
echo ============================================================
echo   바른손 CRM 대시보드 - 로컬 테스트 모드
echo ------------------------------------------------------------
echo   브라우저 주소 : http://localhost:13000
echo   로그인        : barunson  /  barunson2026
echo   데이터 저장   : %DATA_DIR%
echo   종료          : 이 창에서 Ctrl+C
echo ============================================================
echo.
node crm-platform.js
echo.
echo [서버가 종료되었습니다]
pause
