@echo off
set __N__=NUL
chcp 65001 >%__N__% 2>%__N__%
title [최초 1회] Python 설치 + 환경 구성

echo ============================================
echo   POSCO QR 출석 시스템 - 초기 설치
echo   (최초 1회만 실행하면 됩니다)
echo ============================================
echo.

:: ============ Python 경로 탐색 ============
set PYTHON_EXE=

:: 1) 사용자 설치 경로 직접 확인 (가장 확실한 방법)
for %%V in (313 312 311 310 39) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        goto CHECK_PYTHON
    )
)

:: 2) 전체 사용자 설치 경로
for %%V in (313 312 311 310 39) do (
    if exist "C:\Python%%V\python.exe" (
        set "PYTHON_EXE=C:\Python%%V\python.exe"
        goto CHECK_PYTHON
    )
)
for %%V in (313 312 311 310 39) do (
    if exist "C:\Program Files\Python%%V\python.exe" (
        set "PYTHON_EXE=C:\Program Files\Python%%V\python.exe"
        goto CHECK_PYTHON
    )
)

:: 3) py launcher
py --version >%__N__% 2>%__N__%
if %errorlevel% equ 0 (
    set PYTHON_EXE=py
    goto CHECK_PYTHON
)

:: 4) PATH의 python (WindowsApps 스텁 제외)
for /f "tokens=*" %%P in ('where python 2^>%__N__%') do (
    echo %%P | findstr /I WindowsApps >%__N__% 2>%__N__%
    if errorlevel 1 (
        set "PYTHON_EXE=%%P"
        goto CHECK_PYTHON
    )
)

goto PYTHON_NOT_FOUND

:CHECK_PYTHON
:: 실제로 동작하는지 검증 (Windows Store 스텁 걸러냄)
"%PYTHON_EXE%" -c "import sys; print(sys.version)" >%__N__% 2>%__N__%
if %errorlevel% neq 0 (
    echo [WARNING] %PYTHON_EXE% 가 정상 작동하지 않습니다.
    set PYTHON_EXE=
    goto PYTHON_NOT_FOUND
)
goto PYTHON_FOUND

:PYTHON_NOT_FOUND
goto INSTALL_PYTHON

:PYTHON_FOUND
echo [OK] Python이 이미 설치되어 있습니다: %PYTHON_EXE%
echo.
goto INSTALL_PACKAGES

:INSTALL_PYTHON
echo [INFO] Python이 설치되어 있지 않습니다.
echo        지금 Python 3.12를 자동으로 다운로드하고 설치합니다.
echo.

set PYTHON_URL=https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe
set PYTHON_INSTALLER=python_installer.exe

echo [1/3] Python 설치파일 다운로드 중...
echo       (약 25MB, 잠시 기다려 주세요)

powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_INSTALLER%' -UseBasicParsing }"

if not exist "%PYTHON_INSTALLER%" (
    echo.
    echo [ERROR] 다운로드에 실패했습니다.
    echo         인터넷 연결을 확인하거나 아래 주소에서 직접 다운로드하세요:
    echo         https://www.python.org/downloads/
    pause
    exit /b 1
)

echo        다운로드 완료!
echo.

echo [2/3] Python 설치 중... (1~2분 소요)
echo        설치 완료까지 이 창을 닫지 마세요.
echo.

:: 먼저 사용자 폴더에 설치 시도
"%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_test=0

if %errorlevel% neq 0 (
    echo [WARNING] 자동 설치 실패. 수동 설치를 시작합니다.
    echo          반드시 Add Python to PATH 를 체크하세요!
    echo.
    start /wait "" "%PYTHON_INSTALLER%"
)

del /q "%PYTHON_INSTALLER%" 2>%__N__%

echo        Python 설치 완료!
echo.

:: 설치 후 PATH 새로고침
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312\;%LOCALAPPDATA%\Programs\Python\Python312\Scripts\;%PATH%"

:: 설치 확인
set PYTHON_EXE=
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    goto INSTALL_VERIFY_OK
)
for %%V in (313 312 311 310 39) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        goto INSTALL_VERIFY_OK
    )
)

:: PATH로 재확인
for /f "tokens=*" %%P in ('where python 2^>%__N__%') do (
    echo %%P | findstr /I WindowsApps >%__N__% 2>%__N__%
    if errorlevel 1 (
        set "PYTHON_EXE=%%P"
        goto INSTALL_VERIFY_OK
    )
)

echo.
echo [WARNING] 설치는 완료되었으나 Python을 찾을 수 없습니다.
echo           이 창을 닫고 install.bat 을 다시 실행해 주세요.
pause
exit /b 1

:INSTALL_VERIFY_OK
echo [OK] Python 설치 확인: %PYTHON_EXE%
echo.

:INSTALL_PACKAGES
echo [3/3] 필요 패키지 설치 중...
echo        (flask, apscheduler, qrcode, pyngrok)
echo.

"%PYTHON_EXE%" -m pip install --upgrade pip --trusted-host pypi.org --trusted-host files.pythonhosted.org --quiet 2>%__N__%
"%PYTHON_EXE%" -m pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org flask apscheduler "qrcode[pil]" Pillow pyngrok

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 패키지 설치에 실패했습니다.
    echo         인터넷 연결 또는 프록시 설정을 확인해 주세요.
    pause
    exit /b 1
)

:: 설치 검증
"%PYTHON_EXE%" -c "import flask; print('[OK] flask', flask.__version__)"

echo.
echo ============================================
echo   설치가 모두 완료되었습니다!
echo.
echo   이제 run.bat 을 실행하면 서버가 시작됩니다.
echo ============================================
echo.
pause
