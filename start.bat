@echo off
cd /d "%~dp0"
:loop
echo [%date% %time%] CRM dashboard starting...
node crm-platform.js
echo [%date% %time%] server stopped - restarting in 5s
timeout /t 5 /nobreak >nul
goto loop
