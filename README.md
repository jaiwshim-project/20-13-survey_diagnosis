# 조직진단 AI플랫폼 - 로컬 실행 가이드

대한치과위생사협회 조직진단 웹사이트를 로컬 환경에서 테스트하기 위한 가이드입니다.

## 🚀 빠른 시작 (권장)

### Windows 사용자

1. **run_server.bat 파일 더블클릭**
   - 자동으로 서버가 시작되고 브라우저가 열립니다
   - 간단하고 빠릅니다!

### macOS / Linux 사용자

1. **터미널 열기**
2. **다음 명령어 실행:**
   ```bash
   bash run_server.sh
   ```

---

## 📋 상세 실행 방법

### 방법 1️⃣: Python 스크립트 사용 (모든 OS)

**가장 간단한 방법입니다**

1. **터미널/명령 프롬프트 열기**
   - Windows: `Win + R` → `cmd` 입력
   - macOS: `Command + Space` → `Terminal` 검색
   - Linux: 터미널 열기

2. **website 폴더로 이동:**
   ```bash
   cd "C:\01 Claude-Code\20-13 대한치과위생사 유은미 교수\website"
   ```

3. **서버 실행:**
   ```bash
   python run_server.py
   ```
   또는
   ```bash
   python3 run_server.py
   ```

4. **브라우저에서 접속:**
   ```
   http://localhost:8000
   ```

---

### 방법 2️⃣: 내장 Python HTTP 서버 사용

Python이 설치되어 있다면 추가 스크립트 없이 바로 실행 가능합니다.

1. **website 폴더로 이동**
2. **다음 명령어 실행:**
   ```bash
   python -m http.server 8000
   ```

3. **브라우저에서 접속:**
   ```
   http://localhost:8000
   ```

---

### 방법 3️⃣: Node.js 사용 (선택사항)

Node.js와 `http-server`가 설치되어 있다면:

1. **http-server 설치** (처음 1회만):
   ```bash
   npm install -g http-server
   ```

2. **website 폴더로 이동**

3. **서버 실행:**
   ```bash
   http-server
   ```

4. **브라우저에서 접속:**
   ```
   http://127.0.0.1:8080
   ```

---

## 📍 주요 페이지 접속 주소

서버 실행 후 다음 주소들로 접속할 수 있습니다:

| 페이지 | URL |
|--------|-----|
| 메인 페이지 (대한치과위생사협회 진단 메뉴) | `http://localhost:8000/` |
| 위생사협회 소개 | `http://localhost:8000/pages/kdha.html` |
| 설문 선택 | `http://localhost:8000/pages/survey.html` |
| 회장단 사전설문 | `http://localhost:8000/pages/pre-survey-executive.html` |
| 직원 사전설문 | `http://localhost:8000/pages/pre-survey-employee.html` |
| 진단시스템 | `http://localhost:8000/pages/diagnosis.html` |
| 진단영역 | `http://localhost:8000/pages/areas.html` |

---

## ❓ 문제 해결

### ❌ "Python을 찾을 수 없습니다"

**해결책:**
1. Python이 설치되어 있는지 확인
   - Windows: `python --version`
   - macOS/Linux: `python3 --version`

2. Python이 없다면 설치:
   - https://www.python.org/downloads/
   - macOS: `brew install python3`
   - Linux: `sudo apt-get install python3`

### ❌ "포트 8000이 이미 사용 중입니다"

**해결책:**

다른 포트 사용:
```bash
python -m http.server 8001
```
그 다음 `http://localhost:8001` 로 접속

---

### ❌ 이미지나 CSS가 로드되지 않습니다

**해결책:**
1. 브라우저 캐시 삭제: `Ctrl + Shift + Del` (또는 `Cmd + Shift + Del`)
2. 강제 새로고침: `Ctrl + F5` (또는 `Cmd + Shift + R`)
3. 서버 재시작:
   - 터미널에서 `Ctrl + C` 입력 후 다시 실행

---

### ❌ 포트 접근 권한 오류

**해결책:**

macOS/Linux에서 포트 8000 대신 다른 포트 사용:
```bash
python3 -m http.server 8080
```

---

## 🔧 고급 설정

### 자동 캐시 방지 설정

`run_server.py`는 개발 중 캐시 문제를 방지하도록 자동으로 설정됩니다:
- `Cache-Control: no-store, no-cache, must-revalidate`

### 로그 확인

서버 실행 중 모든 요청이 터미널에 표시됩니다:
```
[2026-03-24 10:30:45] GET / HTTP/1.1" 200 -
[2026-03-24 10:30:46] GET /css/style.css HTTP/1.1" 200 -
```

---

## 📌 팁

✅ **개발 중 자주 사용할 것이라면:**
- Windows: `run_server.bat`를 시작 메뉴에 고정
- macOS: `run_server.sh`를 응용프로그램 폴더에 링크

✅ **여러 포트에서 동시 실행:**
```bash
# 터미널 1
python -m http.server 8000

# 터미널 2
python -m http.server 8001
```

✅ **공유 네트워크에서 접속:**
```bash
# 모든 인터페이스에서 수신
python -m http.server 8000 --bind 0.0.0.0
```

---

## 📞 지원

문제가 발생하면:

1. **에러 메시지** 확인 및 복사
2. **터미널 로그** 확인
3. **포트** 변경 후 재시도
4. **브라우저 캐시** 삭제 후 재시도

---

## 📄 파일 설명

- `run_server.py` - Python 웹서버 스크립트 (모든 OS 지원)
- `run_server.bat` - Windows 배치 파일 (더블클릭으로 실행)
- `run_server.sh` - macOS/Linux 쉘 스크립트
- `README.md` - 이 문서

---

**행운을 빕니다! 🚀**
