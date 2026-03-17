@echo off
set __N__=NUL
chcp 65001 >%__N__% 2>%__N__%
title POSCO QR Attendance Server

echo ============================================
echo   POSCO QR 출석 관리 서버
echo   (광양제철소 행정섭외그룹)
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
echo.
echo [ERROR] Python을 찾을 수 없습니다.
echo         install.bat 을 먼저 실행하여 설치해 주세요.
echo.
pause
exit /b 1

:PYTHON_FOUND
echo [OK] Python 확인: %PYTHON_EXE%

:: 패키지 확인
"%PYTHON_EXE%" -c "import flask" >%__N__% 2>%__N__%
if %errorlevel% neq 0 (
    echo [INFO] 필요 패키지가 없습니다. 설치 중...
    "%PYTHON_EXE%" -m pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org flask apscheduler "qrcode[pil]" Pillow pyngrok
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] 패키지 설치에 실패했습니다.
        pause
        exit /b 1
    )
)

:: 서버 시작
echo.
echo [OK] 서버를 시작합니다...
echo   관리자 페이지 : http://localhost:5000
echo   종료          : Ctrl+C
echo.

start /b cmd /c "timeout /t 3 /nobreak >%__N__% 2>%__N__% && start http://localhost:5000"

"%PYTHON_EXE%" app.py

echo.
echo ================================================
echo   서버가 비정상 종료되었습니다.
echo   위 에러 메시지를 확인해 주세요.
echo ================================================
pause
