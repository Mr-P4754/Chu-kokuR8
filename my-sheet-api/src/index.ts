import { Hono } from 'hono';
import { cors } from 'hono/cors';
import { getAuthSpreadsheet, withRetry } from './sheets';

type Bindings = {
  GOOGLE_SERVICE_ACCOUNT_EMAIL: string;
  GOOGLE_PRIVATE_KEY: string;
  SPREADSHEET_ID: string;
  GEMINI_API_KEY: string;
  ADMIN_PIN?: string;
};

const app = new Hono<{ Bindings: Bindings }>();

app.use('*', cors());

const SHEET_QUESTIONS = '【意見・質問】分科会';
const SHEET_RECORDS = '【記録】分科会';
const SHEET_SYSTEM = 'System_Status';
const SHEET_SUMMARIES = 'AI_Summaries';

// --- Helper Functions ---
async function verifyAndHealSheets(doc: any) {
  // 簡易的にシートの存在確認と必要ヘッダーの追加を行う
  // 実際の移行コードに合わせて最適化
  let sysSheet = doc.sheetsByTitle[SHEET_SYSTEM];
  if (!sysSheet) {
    sysSheet = await withRetry(() => doc.addSheet({ title: SHEET_SYSTEM, headerValues: ['sheet_name', 'row_index', 'status', 'is_hidden', 'updated_at'] }));
  }
  
  let aiSheet = doc.sheetsByTitle[SHEET_SUMMARIES];
  if (!aiSheet) {
    aiSheet = await withRetry(() => doc.addSheet({ title: SHEET_SUMMARIES, headerValues: ['id', 'session_id', 'type', 'summary_json', 'created_at'] }));
  }
}

function generateUuid() {
  return crypto.randomUUID();
}

// --- Endpoints ---

app.get('/', (c) => c.text('API is running.'));

// 1. GET /api/questions : 質問およびAI要約データの取得
app.get('/api/questions', async (c) => {
  const sessionId = c.req.query('sessionId');
  if (!sessionId) return c.json({ success: false, error: 'sessionId is required' });

  try {
    const doc = await getAuthSpreadsheet(c.env);
    await verifyAndHealSheets(doc);

    // 質問一覧取得
    const qSheet = doc.sheetsByTitle[SHEET_QUESTIONS];
    const sysSheet = doc.sheetsByTitle[SHEET_SYSTEM];
    
    if (!qSheet) return c.json({ success: true, data: { questions: [], aiSummaries: null } });
    
    const [qRows, sysRows] = await Promise.all([
      withRetry(() => qSheet.getRows()),
      sysSheet ? withRetry(() => sysSheet.getRows()) : Promise.resolve([])
    ]);

    const sysMap: Record<string, any> = {};
    sysRows.forEach(row => {
      if (row.get('sheet_name') === SHEET_QUESTIONS) {
        sysMap[row.get('row_index')] = {
          status: row.get('status'),
          is_hidden: String(row.get('is_hidden')).toUpperCase() === 'TRUE'
        };
      }
    });

    const questions = [];
    for (let i = 0; i < qRows.length; i++) {
      const row = qRows[i];
      // フォーム側（タイムスタンプ, 分科会名...）の想定インデックス
      const formSessionStr = String(row.get('分科会を選択してください') || row.get('分科会名') || row._rawData[1] || '');
      const matchSession = formSessionStr.match(/[A-Za-z]\d+/);
      const formSessionId = matchSession ? matchSession[0].toUpperCase() : formSessionStr.toUpperCase();

      if (formSessionId === sessionId.toUpperCase()) {
        const rowIndex = i + 2; // ヘッダー分(+1)とインデックス(+1)
        const sysInfo = sysMap[rowIndex] || { status: '未回答', is_hidden: false };
        const rawTime = row._rawData[0]; // タイムスタンプ
        
        questions.push({
          rowIndex: rowIndex,
          timestamp: rawTime,
          raw_time: new Date(rawTime).getTime(),
          name: row._rawData[2],
          school: row._rawData[3],
          opinion: row._rawData[4],
          question: row._rawData[5],
          status: sysInfo.status,
          is_hidden: sysInfo.is_hidden
        });
      }
    }

    // AI要約取得
    const aiSheet = doc.sheetsByTitle[SHEET_SUMMARIES];
    let latestQuestions = null, latestRecords = null;
    let qTime = 0, rTime = 0;
    
    if (aiSheet) {
      const aiRows = await withRetry(() => aiSheet.getRows());
      for (const row of aiRows) {
        if (String(row.get('session_id')).toUpperCase() === sessionId.toUpperCase()) {
          const type = row.get('type');
          const time = new Date(row.get('created_at')).getTime();
          if (type === 'QUESTIONS' && time > qTime) {
            latestQuestions = row.get('summary_json');
            qTime = time;
          } else if (type === 'RECORDS' && time > rTime) {
            latestRecords = row.get('summary_json');
            rTime = time;
          }
        }
      }
    }

    const safeParse = (str: string) => {
      try { return str ? JSON.parse(str) : null; } catch (e) { return null; }
    };

    return c.json({
      success: true,
      data: {
        questions,
        aiSummaries: {
          questions: safeParse(latestQuestions),
          records: safeParse(latestRecords),
          qTime: qTime
        }
      }
    });

  } catch (err: any) {
    return c.json({ success: false, error: err.message }, 500);
  }
});

// 2. POST /api/questions/status : ステータスの更新処理
app.post('/api/questions/status', async (c) => {
  try {
    const { rowIndex, status, isHidden } = await c.req.json();
    if (!rowIndex) return c.json({ success: false, error: 'rowIndex is required' });

    const doc = await getAuthSpreadsheet(c.env);
    let sysSheet = doc.sheetsByTitle[SHEET_SYSTEM];
    if (!sysSheet) throw new Error('System sheet not found');

    const rows = await withRetry(() => sysSheet.getRows());
    let found = false;
    for (const row of rows) {
      if (row.get('sheet_name') === SHEET_QUESTIONS && String(row.get('row_index')) === String(rowIndex)) {
        row.assign({ status, is_hidden: isHidden, updated_at: new Date().toISOString() });
        await withRetry(() => row.save());
        found = true;
        break;
      }
    }

    if (!found) {
      await withRetry(() => sysSheet.addRow({
        sheet_name: SHEET_QUESTIONS,
        row_index: rowIndex,
        status: status,
        is_hidden: isHidden,
        updated_at: new Date().toISOString()
      }));
    }

    return c.json({ success: true, data: true });
  } catch (err: any) {
    return c.json({ success: false, error: err.message }, 500);
  }
});

// 3. POST /api/gemini/analyze : Gemini APIを利用した論点整理・要約の実行
app.post('/api/gemini/analyze', async (c) => {
  try {
    const { sessionId, analysisType, scope } = await c.req.json();
    if (!sessionId || !analysisType) return c.json({ success: false, error: 'Invalid parameters' });
    
    if (!c.env.GEMINI_API_KEY) throw new Error('GEMINI_API_KEY is not set');

    const doc = await getAuthSpreadsheet(c.env);
    
    let prompt = '';
    let systemInstruction = '';
    let responseSchema = {};

    if (analysisType === 'QUESTIONS') {
      const qSheet = doc.sheetsByTitle[SHEET_QUESTIONS];
      if (!qSheet) throw new Error('Questions sheet not found');
      
      const qRows = await withRetry(() => qSheet.getRows());
      let textData = '';
      
      for (const row of qRows) {
        const formSessionStr = String(row.get('分科会を選択してください') || row.get('分科会名') || row._rawData[1] || '');
        const matchSession = formSessionStr.match(/[A-Za-z]\d+/);
        const formSessionId = matchSession ? matchSession[0].toUpperCase() : formSessionStr.toUpperCase();
        
        if (formSessionId === sessionId.toUpperCase()) {
          const qText = row.get('質問') || row._rawData[5] || '';
          const opText = row.get('意見') || row.get('感想') || row._rawData[4] || '';
          textData += `質問: ${qText} (意見: ${opText})\n`;
        }
      }
      
      if (!textData.trim()) throw new Error('No data to analyze');
      
      systemInstruction = 'あなたは教育研究大会の質疑応答モデレーターです。提供された質問と意見から論点を整理し、スキーマに従って出力してください。';
      prompt = `以下の分科会${sessionId}の質問データを分析してください:\n${textData}`;
      responseSchema = {
        "type": "OBJECT",
        "properties": {
          "overall_summary": { "type": "STRING", "description": "全体傾向(200字)" },
          "topics": {
            "type": "ARRAY",
            "items": {
              "type": "OBJECT",
              "properties": {
                "topic_title": { "type": "STRING", "description": "論点" },
                "count": { "type": "INTEGER" },
                "summary": { "type": "STRING", "description": "概要" },
                "key_questions": { "type": "ARRAY", "items": { "type": "STRING" } }
              },
              "required": ["topic_title", "count", "summary", "key_questions"]
            }
          }
        },
        "required": ["overall_summary", "topics"]
      };
    } else if (analysisType === 'RECORDS') {
      const rSheet = doc.sheetsByTitle[SHEET_RECORDS];
      if (!rSheet) throw new Error('Records sheet not found');
      
      const rRows = await withRetry(() => rSheet.getRows());
      const records = [];
      for (const row of rRows) {
        const formSessionStr = String(row.get('分科会を選択してください') || row.get('分科会名') || row._rawData[1] || '');
        const matchSession = formSessionStr.match(/[A-Za-z]\d+/);
        const formSessionId = matchSession ? matchSession[0].toUpperCase() : formSessionStr.toUpperCase();
        if (formSessionId === sessionId.toUpperCase()) {
          const recordText = row.get('記録') || row.get('文字起こし') || row._rawData[2] || '';
          if (recordText) records.push(recordText);
        }
      }
      
      if (records.length === 0) throw new Error('No records data to analyze');
      
      systemInstruction = 'あなたは教育研究大会の記録係です。分科会の音声文字起こしから要点を整理し、スキーマに従って要約を出力してください。';
      prompt = `以下の分科会${sessionId}の文字起こしデータを要約してください:\n${records.join('\n')}`;
      responseSchema = {
        "type": "OBJECT",
        "properties": {
          "session_summary": { "type": "STRING", "description": "分科会全体の要約(400字)" },
          "key_takeaways": { "type": "ARRAY", "items": { "type": "STRING" } },
          "discussion_points": { "type": "STRING", "description": "今後の課題や議論されたこと" }
        },
        "required": ["session_summary", "key_takeaways", "discussion_points"]
      };
    }

    const payload = {
      "system_instruction": { "parts": [{ "text": systemInstruction }] },
      "contents": [{ "parts": [{ "text": prompt }] }],
      "generationConfig": {
        "response_mime_type": "application/json",
        "response_schema": responseSchema
      }
    };

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${c.env.GEMINI_API_KEY}`;
    
    let res;
    let attempt = 0;
    while(attempt <= 3) {
      res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (res.status === 429 && attempt < 3) {
        attempt++;
        const delay = (1000 * Math.pow(2, attempt)) + (Math.random() * 1000);
        await new Promise(r => setTimeout(r, delay));
        continue;
      }
      break;
    }

    if (!res || !res.ok) {
      const errText = await res?.text();
      throw new Error(`API Error: ${errText}`);
    }

    const json = await res.json() as any;
    let aiResponseText = json.candidates[0].content.parts[0].text;
    aiResponseText = aiResponseText.replace(/^```(?:json)?\n?/i, '').replace(/\n?```$/i, '').trim();

    // データベース (スプレッドシート) へ保存
    let aiSheet = doc.sheetsByTitle[SHEET_SUMMARIES];
    if (!aiSheet) throw new Error('AI_Summaries sheet not found');
    
    await withRetry(() => aiSheet.addRow({
      id: generateUuid(),
      session_id: sessionId,
      type: analysisType,
      summary_json: aiResponseText,
      created_at: new Date().toISOString()
    }));

    return c.json({ success: true, data: aiResponseText });
  } catch (err: any) {
    return c.json({ success: false, error: err.message }, 500);
  }
});

// 4. POST /api/admin/verify : 管理者PINの認証
app.post('/api/admin/verify', async (c) => {
  try {
    const { pin } = await c.req.json();
    const correctPin = c.env.ADMIN_PIN || '1204';
    if (pin === correctPin) {
      return c.json({ success: true, data: true });
    }
    return c.json({ success: true, data: false });
  } catch (err: any) {
    return c.json({ success: false, error: err.message }, 500);
  }
});

export default app;
