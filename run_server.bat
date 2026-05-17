@echo off
REM Windows용 로컬 웹서버 실행 배치 파일
REM 이 파일을 더블클릭하면 자동으로 서버가 시작됩니다

REM UTF-8 인코딩 설정
chcp 65001 > nul

cls
echo.
echo ========================================================================
echo  조직진단 AI플랫폼 - 로컬 웹서버
echo ========================================================================
echo.
echo 🚀 웹서버를 시작합니다...
echo.

REM Python 설치 확인
where python > nul 2>&1
if errorlevel 1 (
    where python3 > nul 2>&1
    if errorlevel 1 (
        echo ❌ Python이 설치되어 있지 않습니다.
        echo.
        echo 다음 중 하나를 선택하세요:
        echo 1. Python 설치: https://www.python.org/downloads/
        echo 2. 또는 이 명령어를 터미널에서 직접 실행:
        echo    python -m http.server 8000
        echo.
        pause
        exit /b 1
    )
    set PYTHON=python3
) else (
    set PYTHON=python
)

REM 로컬 경로 설정
cd /d "%~dp0"

REM Python 서버 실행
echo 📍 웹 루트: %cd%
echo 🌐 접속 주소: http://localhost:8000
echo ⏸️  종료: 이 창을 닫으면 됩니다 (Ctrl+C 또는 창 닫기)
echo.
echo ========================================================================
echo.

%PYTHON% run_server.py

if errorlevel 1 (
    echo.
    echo ❌ 서버 시작 중 오류가 발생했습니다.
    echo.
    echo 대안: 터미널을 열고 다음 명령어를 입력하세요:
    echo cd "%cd%"
    echo python -m http.server 8000
    echo.
    pause
)

pause
