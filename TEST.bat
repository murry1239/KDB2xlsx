@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo INSTALL.bat 을 먼저 실행하세요.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m unittest discover -s tests -v
set "RESULT=%ERRORLEVEL%"
pause
exit /b %RESULT%
