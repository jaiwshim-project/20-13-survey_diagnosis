// Vercel Serverless Function: Save workshop data
import { createClient } from '@supabase/supabase-js';

const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

export default async function handler(req, res) {
  // CORS 헤더 설정
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // 환경 변수 확인
    if (!supabaseUrl || !supabaseKey) {
      console.error('[ERROR] Missing Supabase credentials');
      return res.status(500).json({ error: 'Server configuration error' });
    }

    const supabase = createClient(supabaseUrl, supabaseKey);

    const { teamNumber, data } = req.body;

    if (!teamNumber || !data) {
      return res.status(400).json({ error: 'Missing teamNumber or data' });
    }

    // Upsert (있으면 업데이트, 없으면 삽입)
    const { data: result, error } = await supabase
      .from('workshop_records')
      .upsert({
        team_number: teamNumber,
        form_data: data,
        updated_at: new Date().toISOString()
      }, {
        onConflict: 'team_number'
      })
      .select();

    if (error) {
      console.error('[Supabase Error]', error);
      return res.status(500).json({ error: error.message });
    }

    return res.status(200).json({
      success: true,
      message: 'Data saved successfully',
      data: result
    });

  } catch (error) {
    console.error('[Server Error]', error);
    return res.status(500).json({ error: error.message });
  }
}
