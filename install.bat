@echo off
chcp 437 >nul 2>&1
setlocal enabledelayedexpansion

if "%~1" neq "__inner__" (
    cmd /c ""%~f0" __inner__"
    pause
    exit /b
)

echo ============================================================
echo   HWPX Writer MCP - Install Script
echo ============================================================
echo.

set "INSTALL_DIR=%~dp0"
if "%INSTALL_DIR:~-1%"=="\" set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"
echo [INFO] Install path: %INSTALL_DIR%
echo.

:: -----------------------------------------------------------
:: 1. Check Python
:: -----------------------------------------------------------
echo [1/4] Checking Python...

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] Python is not installed.
    goto :install_python
)

set "PY_VER="
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set "PY_VER=%%v"

if not defined PY_VER (
    echo [WARNING] Cannot detect Python version.
    echo           Please disable python.exe in App Execution Aliases.
    goto :install_python
)

echo [INFO] Python %PY_VER% detected
echo.

set "PY_MAJOR="
set "PY_MINOR="
for /f "tokens=1,2 delims=." %%a in ("%PY_VER%") do (
    set "PY_MAJOR=%%a"
    set "PY_MINOR=%%b"
)

if not defined PY_MAJOR goto :install_python
if not defined PY_MINOR goto :install_python
if !PY_MAJOR! lss 3 goto :install_python
if !PY_MAJOR!==3 if !PY_MINOR! lss 10 goto :install_python

goto :python_ok

:: -----------------------------------------------------------
:: 2. Auto Install Python
:: -----------------------------------------------------------
:install_python
echo ============================================================
echo   Installing Python 3.11.9
echo ============================================================
echo.
echo Downloading Python 3.11.9...
echo.

set "PY_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
set "PY_INSTALLER=%TEMP%\python-3.11.9-amd64.exe"

echo [Download] %PY_URL%
powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_INSTALLER%' }" 2>nul

if not exist "%PY_INSTALLER%" (
    echo.
    echo [ERROR] Failed to download Python.
    echo         Please install manually: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [Installing] Python 3.11.9...
echo.

"%PY_INSTALLER%" /passive InstallAllUsers=0 PrependPath=1 Include_test=0

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Python installation failed.
    echo         Please install manually: https://www.python.org/downloads/
    pause
    exit /b 1
)

set "PATH=%LOCALAPPDATA%\Programs\Python\Python311\Scripts\;%LOCALAPPDATA%\Programs\Python\Python311\;%PATH%"
del "%PY_INSTALLER%" >nul 2>&1
echo [OK] Python 3.11.9 installed!
echo.

:python_ok

:: -----------------------------------------------------------
:: 3. Create virtual environment
:: -----------------------------------------------------------
echo [2/4] Creating virtual environment...

if exist "%INSTALL_DIR%\.venv\Scripts\python.exe" (
    echo [INFO] Virtual environment already exists. Skipping.
) else (
    python -m venv "%INSTALL_DIR%\.venv"
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
)
echo.

:: -----------------------------------------------------------
:: 4. Install packages
:: -----------------------------------------------------------
echo [3/4] Installing packages...

"%INSTALL_DIR%\.venv\Scripts\pip.exe" install -r "%INSTALL_DIR%\requirements.txt" --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install packages.
    pause
    exit /b 1
)
echo [OK] Packages installed.
echo.

:: -----------------------------------------------------------
:: 5. Done
:: -----------------------------------------------------------
echo ============================================================
echo   Installation Complete!
echo ============================================================
echo.
echo Add the following to your Claude Desktop config file:
echo.
echo Config file location:
echo   %%APPDATA%%\Claude\claude_desktop_config.json
echo.
echo ------- Copy below -------
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
echo ------- End of copy -------
echo.
echo Restart Claude Desktop after saving the config.
echo.
