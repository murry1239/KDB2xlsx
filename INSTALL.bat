@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PY_CMD="
where py >nul 2>nul
if not errorlevel 1 (
  py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)" >nul 2>nul
  if not errorlevel 1 set "PY_CMD=py -3"
)
if not defined PY_CMD (
  where python >nul 2>nul
  if not errorlevel 1 (
    python -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)" >nul 2>nul
    if not errorlevel 1 set "PY_CMD=python"
  )
)
if not defined PY_CMD goto no_python

echo [1/3] 가상환경 만드는 중...
%PY_CMD% -m venv ".venv"
if errorlevel 1 goto failed
set "VENV_PY=%CD%\.venv\Scripts\python.exe"
if not exist "%VENV_PY%" goto failed

echo [2/3] pip 업데이트 중...
"%VENV_PY%" -m pip install --upgrade pip
if errorlevel 1 goto failed

echo [3/3] 필요한 패키지 설치 중...
"%VENV_PY%" -m pip install -r "requirements.txt"
if errorlevel 1 goto failed

echo.
echo 설치가 끝났습니다. RUN.bat 을 더블클릭하여 실행하세요.
goto finish

:no_python
echo.
echo Python 3.11 이상을 찾지 못했습니다.
echo 64비트 Python 을 설치(Python Launcher 포함)한 뒤 다시 실행하세요.
goto finish

:failed
echo.
echo 설치에 실패했습니다. DIAGNOSE.bat 을 실행하고 결과를 저장해 전달해 주세요.

:finish
echo.
pause
endlocal
