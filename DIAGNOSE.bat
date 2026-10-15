@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo 원시자료 정리봇 진단
echo 작업 폴더: %CD%
echo 로그 폴더: %LOCALAPPDATA%\원시자료정리봇\logs
where py 2>nul
where python 2>nul
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" --version
  ".venv\Scripts\python.exe" -m pip show openpyxl
  ".venv\Scripts\python.exe" -c "import tkinter, openpyxl, engine; print('화면/엑셀 엔진 정상:', engine.VERSION)"
) else (
  echo 가상환경이 없습니다. INSTALL.bat 을 실행하세요.
)
pause
endlocal
