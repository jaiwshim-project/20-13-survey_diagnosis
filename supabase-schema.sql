-- =========================================================
-- Supabase 테이블 생성 SQL
-- Supabase > SQL Editor에서 전체 코드를 복사하여 실행하세요.
-- =========================================================

-- ========================
-- 1. 사전설문 회장단 (10문항)
-- ========================
CREATE TABLE pre_survey_executive (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  submitted_at TIMESTAMPTZ DEFAULT NOW(),
  q1 TEXT,
  q2 TEXT,
  q3 TEXT,
  q4 TEXT,
  q5 TEXT,
  q6 TEXT,
  q7 TEXT,
  q8 TEXT,
  q9 TEXT,
  q10 INTEGER  -- 5점 척도
);

-- ========================
-- 2. 사전설문 직원 (10문항)
-- ========================
CREATE TABLE pre_survey_employee (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  submitted_at TIMESTAMPTZ DEFAULT NOW(),
  q1 TEXT,
  q2 TEXT,
  q3 TEXT,
  q4 TEXT,
  q5 TEXT,
  q6 TEXT,
  q7 TEXT,
  q8 TEXT,
  q9 TEXT,
  q10 INTEGER  -- 5점 척도
);

-- ========================
-- 3. 본진단 회장단 (30문항)
-- ========================
CREATE TABLE executive_survey (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  submitted_at TIMESTAMPTZ DEFAULT NOW(),
  q1 TEXT,
  q2 TEXT,
  q3 TEXT,
  q4 TEXT,
  q5 INTEGER,   -- 7점 척도
  q6 TEXT,
  q7 TEXT,
  q8 TEXT,
  q9 TEXT,
  q10 INTEGER,  -- 7점 척도
  q11 TEXT,
  q12 TEXT,
  q13 TEXT,
  q14 TEXT,
  q15 INTEGER,  -- 7점 척도
  q16 TEXT,
  q17 TEXT,
  q18 TEXT,
  q19 TEXT,
  q20 INTEGER,  -- 7점 척도
  q21 TEXT,
  q22 TEXT,
  q23 TEXT,
  q24 TEXT,
  q25 INTEGER,  -- 7점 척도
  q26 TEXT,
  q27 TEXT,
  q28 TEXT,
  q29 TEXT,
  q30 TEXT
);

-- ========================
-- 4. 본진단 직원 (30문항)
-- ========================
CREATE TABLE employee_survey (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  submitted_at TIMESTAMPTZ DEFAULT NOW(),
  q1 TEXT,
  q2 TEXT,
  q3 TEXT,
  q4 TEXT,
  q5 INTEGER,   -- 5점 척도
  q6 TEXT,
  q7 TEXT,
  q8 TEXT,
  q9 TEXT,
  q10 INTEGER,  -- 5점 척도
  q11 TEXT,
  q12 TEXT,
  q13 TEXT,
  q14 TEXT,
  q15 INTEGER,  -- 5점 척도
  q16 TEXT,
  q17 TEXT,
  q18 TEXT,
  q19 TEXT,
  q20 INTEGER,  -- 5점 척도
  q21 TEXT,
  q22 TEXT,
  q23 TEXT,
  q24 TEXT,
  q25 INTEGER,  -- 5점 척도
  q26 TEXT,
  q27 TEXT,
  q28 TEXT,
  q29 TEXT,
  q30 TEXT
);

-- =========================================================
-- Row Level Security (RLS) 설정
-- 익명 사용자는 INSERT만 가능 (읽기 불가)
-- =========================================================

-- 1. RLS 활성화
ALTER TABLE pre_survey_executive ENABLE ROW LEVEL SECURITY;
ALTER TABLE pre_survey_employee ENABLE ROW LEVEL SECURITY;
ALTER TABLE executive_survey ENABLE ROW LEVEL SECURITY;
ALTER TABLE employee_survey ENABLE ROW LEVEL SECURITY;

-- 2. 익명 사용자 INSERT 정책
CREATE POLICY "allow_insert_anon" ON pre_survey_executive
  FOR INSERT TO anon
  WITH CHECK (TRUE);

CREATE POLICY "allow_insert_anon" ON pre_survey_employee
  FOR INSERT TO anon
  WITH CHECK (TRUE);

CREATE POLICY "allow_insert_anon" ON executive_survey
  FOR INSERT TO anon
  WITH CHECK (TRUE);

CREATE POLICY "allow_insert_anon" ON employee_survey
  FOR INSERT TO anon
  WITH CHECK (TRUE);

-- =========================================================
-- 완료!
-- 이제 다음 단계를 수행하세요:
-- 1. js/supabase-config.js에 Supabase URL과 anon key 입력
-- 2. 4개 설문 페이지 업데이트 완료 후 Vercel에 배포
-- =========================================================
