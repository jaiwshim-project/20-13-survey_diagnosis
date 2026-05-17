#!/usr/bin/env python3
"""
로컬 웹서버 실행 스크립트
조직진단 AI플랫폼 사이트를 로컬에서 테스트하기 위한 간단한 HTTP 서버
"""

import http.server
import socketserver
import os
import webbrowser
from pathlib import Path
import sys

# 포트 설정
PORT = 8000
HOST = 'localhost'

# 현재 디렉토리를 웹 루트로 설정
WEB_ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(WEB_ROOT)

class MyHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # 캐시 방지 (개발 중에는 필수)
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Expires', '0')
        super().end_headers()

    def log_message(self, format, *args):
        # 로그 출력 향상
        print(f"[{self.log_date_time_string()}] {format % args}")

def main():
    """메인 함수"""
    url = f"http://{HOST}:{PORT}"

    print("=" * 70)
    print("🚀 조직진단 AI플랫폼 - 로컬 웹서버 시작")
    print("=" * 70)
    print(f"\n📍 웹 루트: {WEB_ROOT}")
    print(f"🌐 접속 주소: {url}")
    print(f"⏸️  종료: Ctrl+C 를 누르세요\n")

    # 서버 생성
    try:
        with socketserver.TCPServer((HOST, PORT), MyHTTPRequestHandler) as httpd:
            print("=" * 70)
            print("✅ 서버가 시작되었습니다!")
            print("=" * 70)

            # 자동으로 브라우저 열기 (선택적)
            try:
                print(f"\n🌍 브라우저를 열고 있습니다...\n")
                webbrowser.open(url)
            except:
                print(f"\n💡 브라우저를 수동으로 열어주세요: {url}\n")

            # 서버 실행
            httpd.serve_forever()

    except OSError as e:
        if e.errno == 48 or e.errno == 98:  # Port already in use
            print(f"\n❌ 오류: 포트 {PORT}가 이미 사용 중입니다.")
            print(f"   다른 서버를 종료하거나 포트를 변경해주세요.")
            sys.exit(1)
        else:
            print(f"\n❌ 오류: {e}")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n⏹️  서버를 종료합니다...")
        print("감사합니다! 👋")
        sys.exit(0)

if __name__ == "__main__":
    main()
