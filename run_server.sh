#!/bin/bash

# macOS/Linux용 로컬 웹서버 실행 쉘 스크립트
# 터미널에서 ./run_server.sh 또는 bash run_server.sh 로 실행

# 색상 설정
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 현재 디렉토리
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

echo ""
echo "========================================================================"
echo -e "${BLUE}🚀 조직진단 AI플랫폼 - 로컬 웹서버${NC}"
echo "========================================================================"
echo ""

# 스크립트 디렉토리로 이동
cd "$SCRIPT_DIR"

# Python 버전 확인
if command -v python3 &> /dev/null; then
    PYTHON=python3
elif command -v python &> /dev/null; then
    PYTHON=python
else
    echo -e "${RED}❌ Python이 설치되어 있지 않습니다.${NC}"
    echo ""
    echo "설치 방법:"
    echo "  macOS: brew install python3"
    echo "  Linux: sudo apt-get install python3"
    echo ""
    exit 1
fi

echo -e "${GREEN}✅ Python 버전:${NC}"
$PYTHON --version
echo ""

echo -e "${BLUE}📍 웹 루트:${NC} $SCRIPT_DIR"
echo -e "${BLUE}🌐 접속 주소:${NC} http://localhost:8000"
echo -e "${BLUE}⏸️  종료:${NC} Ctrl+C를 누르세요"
echo ""
echo "========================================================================"
echo ""

# Python 서버 실행
$PYTHON run_server.py

if [ $? -ne 0 ]; then
    echo ""
    echo -e "${RED}❌ 서버 시작 중 오류가 발생했습니다.${NC}"
    echo ""
    echo "대안: 다음 명령어를 터미널에서 입력하세요:"
    echo "  cd $SCRIPT_DIR"
    echo "  python3 -m http.server 8000"
    echo ""
    exit 1
fi
