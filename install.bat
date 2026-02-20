@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

echo ============================================================
echo   HWPX Writer MCP - 설치 스크립트
echo ============================================================
echo.

:: -----------------------------------------------------------
:: 1. 현재 폴더 확인
:: -----------------------------------------------------------
set "INSTALL_DIR=%~dp0"
:: 끝의 \ 제거
if "%INSTALL_DIR:~-1%"=="\" set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"
echo [정보] 설치 경로: %INSTALL_DIR%
echo.

:: -----------------------------------------------------------
:: 2. Python 설치 여부 확인
:: -----------------------------------------------------------
echo [1/4] Python 확인 중...

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [경고] Python이 설치되어 있지 않습니다.
    echo.
    goto :install_python
)

:: 버전 확인 (3.10 이상 필요)
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set "PY_VER=%%v"
echo [정보] Python %PY_VER% 감지됨
echo.

:: 메이저.마이너 추출
for /f "tokens=1,2 delims=." %%a in ("%PY_VER%") do (
    set "PY_MAJOR=%%a"
    set "PY_MINOR=%%b"
)

if %PY_MAJOR% lss 3 goto :install_python
if %PY_MAJOR%==3 if %PY_MINOR% lss 10 goto :install_python

goto :python_ok

:: -----------------------------------------------------------
:: 3. Python 자동 설치
:: -----------------------------------------------------------
:install_python
echo ============================================================
echo   Python 3.11.9 설치
echo ============================================================
echo.
echo Python 3.11.9 (안정 버전)을 다운로드합니다...
echo.

set "PY_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
set "PY_INSTALLER=%TEMP%\python-3.11.9-amd64.exe"

:: 다운로드
echo [다운로드] %PY_URL%
powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_INSTALLER%' }" 2>nul

if not exist "%PY_INSTALLER%" (
    echo.
    echo [오류] Python 다운로드에 실패했습니다.
    echo        수동으로 설치해주세요: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [설치] Python 3.11.9 설치 중... (PATH 자동 추가)
echo        설치 창이 뜨면 완료될 때까지 기다려주세요.
echo.

"%PY_INSTALLER%" /passive InstallAllUsers=0 PrependPath=1 Include_test=0

if %errorlevel% neq 0 (
    echo.
    echo [오류] Python 설치에 실패했습니다.
    echo        수동으로 설치해주세요: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: PATH 갱신
set "PATH=%LOCALAPPDATA%\Programs\Python\Python311\Scripts\;%LOCALAPPDATA%\Programs\Python\Python311\;%PATH%"

del "%PY_INSTALLER%" >nul 2>&1

echo [완료] Python 3.11.9 설치 완료!
echo.

:python_ok

:: -----------------------------------------------------------
:: 4. 가상환경 생성
:: -----------------------------------------------------------
echo [2/4] 가상환경 생성 중...

if exist "%INSTALL_DIR%\.venv\Scripts\python.exe" (
    echo [정보] 기존 가상환경이 있습니다. 건너뜁니다.
) else (
    python -m venv "%INSTALL_DIR%\.venv"
    if %errorlevel% neq 0 (
        echo [오류] 가상환경 생성에 실패했습니다.
        pause
        exit /b 1
    )
    echo [완료] 가상환경 생성 완료
)
echo.

:: -----------------------------------------------------------
:: 5. 패키지 설치
:: -----------------------------------------------------------
echo [3/4] 패키지 설치 중...

"%INSTALL_DIR%\.venv\Scripts\pip.exe" install -r "%INSTALL_DIR%\requirements.txt" --quiet
if %errorlevel% neq 0 (
    echo [오류] 패키지 설치에 실패했습니다.
    pause
    exit /b 1
)
echo [완료] 패키지 설치 완료
echo.

:: -----------------------------------------------------------
:: 6. 설치 완료 안내
:: -----------------------------------------------------------
echo ============================================================
echo   설치 완료!
echo ============================================================
echo.
echo Claude Desktop 설정 파일에 아래 내용을 추가하세요.
echo.
echo 설정 파일 위치:
echo   %%APPDATA%%\Claude\claude_desktop_config.json
echo.
echo ------- 아래를 복사하세요 -------
echo.
echo {
echo   "mcpServers": {
echo     "hwpx-writer": {
echo       "command": "%INSTALL_DIR%\.venv\Scripts\python.exe",
echo       "args": ["%INSTALL_DIR%\server.py"]
echo     }
echo   }
echo }
echo.
echo ------- 여기까지 -------
echo.
echo 설정 후 Claude Desktop을 재시작하면 사용할 수 있습니다.
echo.
pause
