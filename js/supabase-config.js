// =========================================================
// Supabase 설정
//
// 작동 방법:
// 1. https://supabase.com 로그인
// 2. 프로젝트 선택 > Settings > API
// 3. 아래 두 값을 복사하여 붙여넣으세요:
//    - Project URL (https://xxxxx.supabase.co)
//    - anon (public) key
// =========================================================

const SUPABASE_URL = 'https://sqfuqnxlafcilsookmqm.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNxZnVxbnhsYWZjaWxzb29rbXFtIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQyODI4NTMsImV4cCI6MjA4OTg1ODg1M30.QG99xzmjBAIPMXQaC_vpzwKdlIwp4nwUDkqdy2sNz54';

// Supabase 클라이언트 초기화
// (Supabase CDN은 각 설문 HTML 파일의 <head>에 포함됨)
