-- 워크숍 기록 테이블 생성
CREATE TABLE IF NOT EXISTS workshop_records (
  id SERIAL PRIMARY KEY,
  team_number INTEGER NOT NULL UNIQUE,
  form_data JSONB NOT NULL,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 인덱스 생성 (team_number로 빠른 조회)
CREATE INDEX IF NOT EXISTS idx_workshop_records_team_number ON workshop_records(team_number);

-- RLS (Row Level Security) 설정 - 모든 사용자가 읽기/쓰기 가능
ALTER TABLE workshop_records ENABLE ROW LEVEL SECURITY;

CREATE POLICY "워크숍 기록 읽기 허용"
  ON workshop_records FOR SELECT
  USING (true);

CREATE POLICY "워크숍 기록 쓰기 허용"
  ON workshop_records FOR INSERT
  WITH CHECK (true);

CREATE POLICY "워크숍 기록 업데이트 허용"
  ON workshop_records FOR UPDATE
  USING (true);

-- 테이블 코멘트
COMMENT ON TABLE workshop_records IS '대한치과위생사협회 비전워크숍 팀별 실습 기록';
COMMENT ON COLUMN workshop_records.team_number IS '팀 번호 (1~4)';
COMMENT ON COLUMN workshop_records.form_data IS '실습 기록지 전체 데이터 (JSON)';
