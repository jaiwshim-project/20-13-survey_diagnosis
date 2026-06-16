// Vercel Serverless Function: Load workshop data
import { createClient } from '@supabase/supabase-js';

const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

export default async function handler(req, res) {
  // CORS 헤더 설정
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // 환경 변수 확인
    if (!supabaseUrl || !supabaseKey) {
      console.error('[ERROR] Missing Supabase credentials');
      return res.status(500).json({ error: 'Server configuration error' });
    }

    const supabase = createClient(supabaseUrl, supabaseKey);

    const { teamNumber } = req.query;

    if (!teamNumber) {
      return res.status(400).json({ error: 'Missing teamNumber parameter' });
    }

    // 데이터 조회
    const { data, error } = await supabase
      .from('workshop_records')
      .select('*')
      .eq('team_number', parseInt(teamNumber))
      .single();

    if (error) {
      // 데이터가 없는 경우 빈 응답 반환
      if (error.code === 'PGRST116') {
        return res.status(200).json({ success: true, data: null });
      }
      console.error('[Supabase Error]', error);
      return res.status(500).json({ error: error.message });
    }

    return res.status(200).json({
      success: true,
      data: data
    });

  } catch (error) {
    console.error('[Server Error]', error);
    return res.status(500).json({ error: error.message });
  }
}
