import { Hono } from 'hono';
import { cors } from 'hono/cors';
import { getAuthSpreadsheet, withRetry } from './sheets';

type Bindings = {
  GOOGLE_SERVICE_ACCOUNT_EMAIL: string;
  GOOGLE_PRIVATE_KEY: string;
  SPREADSHEET_ID: string;
  GEMINI_API_KEY: string;
  GEMINI_MODEL?: string;
  ADMIN_PIN?: string;
};

const app = new Hono<{ Bindings: Bindings }>();

app.use('*', cors());

// 分科会IDと【意見・質問】シート名の対応マップ
export const SESSION_SHEET_MAP: Record<string, string> = {
  'A-1': '【意見・質問】A-1分科会',
  'A-2': '【意見・質問】A-2分科会',
  'A-3': '【意見・質問】A-3分科会',
  'B-1①': '【意見・質問】B-1分科会①',
  'B-1②': '【意見・質問】B-1分科会②',
  'B-2①': '【意見・質問】B-2分科会①',
  'B-2②': '【意見・質問】B-2分科会②',
  'B-3①': '【意見・質問】B-3分科会①',
  'B-3②': '【意見・質問】B-3分科会②',
  'B-4①': '【意見・質問】B-4分科会①',
  'B-4②': '【意見・質問】B-4分科会②',
};

const QUESTION_SHEET_HEADERS = [
  'タイムスタンプ',
  '名前を記入してください。',
  '所属校を記入してください。',
  '実践報告に対する感想や意見を入力してください。',
  '実践報告に対する質問を入力してください。',
];

const SHEET_RECORDS = '【記録】分科会';
const SHEET_SYSTEM = 'System_Status';
const SHEET_SUMMARIES = 'AI_Summaries';

// --- ヘルパー関数 ---

/**
 * 文字列から正規化された分科会ID（例: 'A-1', 'B-1①'）を抽出する
 */
function normalizeSessionId(rawText: string): string {
  if (!rawText) return '';
  const trimmed = String(rawText).trim();
  if (SESSION_SHEET_MAP[trimmed]) return trimmed;

  // 全角英数字・ハイフンを半角に統一し、「分科会」という語を除去してパターン抽出
  const normalized = trimmed
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[ー－―‐]/g, '-')
    .replace(/分科会/g, '');

  const match = normalized.match(/([AB])-?(\d)\s*([①②])?/i);
  if (match) {
    const group = match[1].toUpperCase();
    const num = match[2];
    const sub = match[3] || '';
    const candidate = `${group}-${num}${sub}`;
    if (SESSION_SHEET_MAP[candidate]) return candidate;
    return candidate;
  }
  return trimmed.toUpperCase();
}

/**
 * 分科会IDに対応する【意見・質問】シート名を取得する
 */
function getQuestionSheetTitle(sessionId: string): string {
  const normalizedId = normalizeSessionId(sessionId);
  return SESSION_SHEET_MAP[normalizedId] || `【意見・質問】${normalizedId}分科会`;
}

/**
 * 質問シートの1行から各フィールドを抽出する（5列構成・旧6列構成の両対応）
 */
function extractQuestionRowData(row: any) {
  const rawData: any[] = row._rawData || [];
  const hasSessionColumn =
    row.get('参加した分科会') !== undefined ||
    row.get('分科会を選択してください') !== undefined ||
    row.get('分科会名') !== undefined ||
    rawData.length >= 6;

  const timestamp = row.get('タイムスタンプ') ?? rawData[0] ?? '';
  const name =
    row.get('名前を記入してください。') ??
    row.get('名前') ??
    (hasSessionColumn ? rawData[2] : rawData[1]) ??
    '';
  const school =
    row.get('所属校を記入してください。') ??
    row.get('所属校') ??
    (hasSessionColumn ? rawData[3] : rawData[2]) ??
    '';
  const opinion =
    row.get('実践報告に対する感想や意見を入力してください。') ??
    row.get('感想・意見') ??
    row.get('意見') ??
    (hasSessionColumn ? rawData[4] : rawData[3]) ??
    '';
  const question =
    row.get('実践報告に対する質問を入力してください。') ??
    row.get('質問') ??
    (hasSessionColumn ? rawData[5] : rawData[4]) ??
    '';

  return { timestamp, name, school, opinion, question };
}

/**
 * 必須システムシートおよび対象分科会シートの存在確認と自己修復を行う
 */
async function verifyAndHealSheets(doc: any, targetQuestionSheetTitle?: string) {
  let sysSheet = doc.sheetsByTitle[SHEET_SYSTEM];
  if (!sysSheet) {
    sysSheet = await withRetry(() =>
      doc.addSheet({
        title: SHEET_SYSTEM,
        headerValues: ['sheet_name', 'row_index', 'status', 'is_hidden', 'updated_at'],
      })
    );
  }

  let aiSheet = doc.sheetsByTitle[SHEET_SUMMARIES];
  if (!aiSheet) {
    aiSheet = await withRetry(() =>
      doc.addSheet({
        title: SHEET_SUMMARIES,
        headerValues: ['id', 'session_id', 'type', 'summary_json', 'created_at'],
      })
    );
  }

  if (targetQuestionSheetTitle && !doc.sheetsByTitle[targetQuestionSheetTitle]) {
    await withRetry(() =>
      doc.addSheet({
        title: targetQuestionSheetTitle,
        headerValues: QUESTION_SHEET_HEADERS,
      })
    );
  }
}

function generateUuid() {
  return crypto.randomUUID();
}

// --- エンドポイント ---

app.get('/', (c) => c.text('API is running.'));

// 1. GET /api/questions : 質問およびAI要約データの取得
app.get('/api/questions', async (c) => {
  const rawSessionId = c.req.query('sessionId');
  if (!rawSessionId) return c.json({ success: false, error: 'sessionId is required' });

  const sessionId = normalizeSessionId(rawSessionId);
  const questionSheetTitle = getQuestionSheetTitle(sessionId);

  try {
    const doc = await getAuthSpreadsheet(c.env);
    await verifyAndHealSheets(doc, questionSheetTitle);

    // 該当分科会の質問シートを取得
    const qSheet = doc.sheetsByTitle[questionSheetTitle];
    const sysSheet = doc.sheetsByTitle[SHEET_SYSTEM];

    if (!qSheet) return c.json({ success: true, data: { questions: [], aiSummaries: null } });

    const [qRows, sysRows] = await Promise.all([
      withRetry(() => qSheet.getRows()),
      sysSheet ? withRetry(() => sysSheet.getRows()) : Promise.resolve([]),
    ]);

    const sysMap: Record<string, any> = {};
    sysRows.forEach((row: any) => {
      if (row.get('sheet_name') === questionSheetTitle) {
        sysMap[row.get('row_index')] = {
          status: row.get('status'),
          is_hidden: String(row.get('is_hidden')).toUpperCase() === 'TRUE',
        };
      }
    });

    const questions = [];
    for (let i = 0; i < qRows.length; i++) {
      const row = qRows[i];
      const { timestamp, name, school, opinion, question } = extractQuestionRowData(row);

      // 空行はスキップ
      if (!timestamp && !name && !opinion && !question) continue;

      const rowIndex = i + 2; // ヘッダー分(+1)とインデックス(+1)
      const sysInfo = sysMap[rowIndex] || { status: '未回答', is_hidden: false };

      questions.push({
        rowIndex: rowIndex,
        timestamp: timestamp,
        raw_time: timestamp ? new Date(timestamp).getTime() : 0,
        name: name,
        school: school,
        opinion: opinion,
        question: question,
        status: sysInfo.status,
        is_hidden: sysInfo.is_hidden,
      });
    }

    // AI要約取得
    const aiSheet = doc.sheetsByTitle[SHEET_SUMMARIES];
    let latestQuestions = null,
      latestRecords = null;
    let qTime = 0,
      rTime = 0;

    if (aiSheet) {
      const aiRows = await withRetry(() => aiSheet.getRows());
      for (const row of aiRows) {
        if (normalizeSessionId(String(row.get('session_id'))) === sessionId) {
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
      try {
        return str ? JSON.parse(str) : null;
      } catch (e) {
        return null;
      }
    };

    return c.json({
      success: true,
      data: {
        questions,
        aiSummaries: {
          questions: safeParse(latestQuestions),
          records: safeParse(latestRecords),
          qTime: qTime,
        },
      },
    });
  } catch (err: any) {
    return c.json({ success: false, error: err.message }, 500);
  }
});

// 2. POST /api/questions/status : ステータスの更新処理
app.post('/api/questions/status', async (c) => {
  try {
    const { sessionId, rowIndex, status, isHidden } = await c.req.json();
    if (!rowIndex) return c.json({ success: false, error: 'rowIndex is required' });

    const targetSheetTitle = sessionId ? getQuestionSheetTitle(sessionId) : '【意見・質問】A-1分科会';

    const doc = await getAuthSpreadsheet(c.env);
    const sysSheet = doc.sheetsByTitle[SHEET_SYSTEM];
    if (!sysSheet) throw new Error('System sheet not found');

    const rows = await withRetry(() => sysSheet.getRows());
    let found = false;
    for (const row of rows) {
      if (row.get('sheet_name') === targetSheetTitle && String(row.get('row_index')) === String(rowIndex)) {
        row.assign({ status, is_hidden: isHidden, updated_at: new Date().toISOString() });
        await withRetry(() => row.save());
        found = true;
        break;
      }
    }

    if (!found) {
      await withRetry(() =>
        sysSheet.addRow({
          sheet_name: targetSheetTitle,
          row_index: rowIndex,
          status: status,
          is_hidden: isHidden,
          updated_at: new Date().toISOString(),
        })
      );
    }

    return c.json({ success: true, data: true });
  } catch (err: any) {
    return c.json({ success: false, error: err.message }, 500);
  }
});

// 3. POST /api/gemini/analyze : Gemini APIを利用した論点整理・要約の実行
app.post('/api/gemini/analyze', async (c) => {
  try {
    const { sessionId: rawSessionId, analysisType, scope } = await c.req.json();
    if (!rawSessionId || !analysisType) return c.json({ success: false, error: 'Invalid parameters' });

    if (!c.env.GEMINI_API_KEY) throw new Error('GEMINI_API_KEY is not set');

    const sessionId = normalizeSessionId(rawSessionId);
    const questionSheetTitle = getQuestionSheetTitle(sessionId);

    const doc = await getAuthSpreadsheet(c.env);

    let prompt = '';
    let systemInstruction = '';
    let responseSchema = {};

    if (analysisType === 'QUESTIONS') {
      const qSheet = doc.sheetsByTitle[questionSheetTitle];
      if (!qSheet) throw new Error(`質問シート（${questionSheetTitle}）が見つかりません`);

      // 未分析データのみ（scope === 'new'）の場合、前回の分析時刻を取得
      let latestAITime = 0;
      if (scope === 'new') {
        const aiSheet = doc.sheetsByTitle[SHEET_SUMMARIES];
        if (aiSheet) {
          const aiRows = await withRetry(() => aiSheet.getRows());
          for (const row of aiRows) {
            if (normalizeSessionId(String(row.get('session_id'))) === sessionId && row.get('type') === 'QUESTIONS') {
              const time = new Date(row.get('created_at')).getTime();
              if (time > latestAITime) latestAITime = time;
            }
          }
        }
      }

      const qRows = await withRetry(() => qSheet.getRows());
      let textData = '';

      for (const row of qRows) {
        const { timestamp, opinion, question } = extractQuestionRowData(row);
        if (!opinion && !question) continue;

        if (scope === 'new' && latestAITime > 0) {
          const rawTime = timestamp ? new Date(timestamp).getTime() : 0;
          if (rawTime <= latestAITime) continue;
        }

        textData += `質問: ${question} (意見: ${opinion})\n`;
      }

      if (!textData.trim()) {
        throw new Error(scope === 'new' ? '未分析の新しいデータがありません' : '分析するデータがありません');
      }

      systemInstruction =
        'あなたは教育研究大会の質疑応答モデレーターです。提供された質問と意見から論点を整理し、スキーマに従って出力してください。';
      prompt = `以下の分科会${sessionId}の質問データを分析してください:\n${textData}`;
      responseSchema = {
        type: 'OBJECT',
        properties: {
          overall_summary: { type: 'STRING', description: '全体傾向(200字)' },
          topics: {
            type: 'ARRAY',
            items: {
              type: 'OBJECT',
              properties: {
                topic_title: { type: 'STRING', description: '論点' },
                count: { type: 'INTEGER' },
                summary: { type: 'STRING', description: '概要' },
                key_questions: { type: 'ARRAY', items: { type: 'STRING' } },
              },
              required: ['topic_title', 'count', 'summary', 'key_questions'],
            },
          },
        },
        required: ['overall_summary', 'topics'],
      };
    } else if (analysisType === 'RECORDS') {
      const rSheet = doc.sheetsByTitle[SHEET_RECORDS];
      if (!rSheet) throw new Error('Records sheet not found');

      const rRows = await withRetry(() => rSheet.getRows());
      const records = [];
      for (const row of rRows) {
        const formSessionStr = String(
          row.get('記録会場を選択してください。') ||
            row.get('記録会場を選択してください') ||
            row.get('参加した分科会') ||
            row.get('分科会を選択してください') ||
            row.get('分科会名') ||
            row._rawData[1] ||
            ''
        );
        const formSessionId = normalizeSessionId(formSessionStr);
        if (formSessionId === sessionId) {
          const recordText =
            row.get('分科会発表の音声文字起こしを入力してください。') ||
            row.get('分科会発表の音声文字起こしを入力してください') ||
            row.get('記録') ||
            row.get('文字起こし') ||
            row._rawData[2] ||
            '';
          if (recordText) records.push(recordText);
        }
      }

      if (records.length === 0) throw new Error('分析する記録データがありません');

      systemInstruction =
        'あなたは教育研究大会の記録係です。分科会の音声文字起こしから要点を整理し、スキーマに従って要約を出力してください。';
      prompt = `以下の分科会${sessionId}の文字起こしデータを要約してください:\n${records.join('\n')}`;
      responseSchema = {
        type: 'OBJECT',
        properties: {
          session_summary: { type: 'STRING', description: '分科会全体の要約(400字)' },
          key_takeaways: { type: 'ARRAY', items: { type: 'STRING' } },
          discussion_points: { type: 'STRING', description: '今後の課題や議論されたこと' },
        },
        required: ['session_summary', 'key_takeaways', 'discussion_points'],
      };
    }

    const payload = {
      system_instruction: { parts: [{ text: systemInstruction }] },
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: {
        response_mime_type: 'application/json',
        response_schema: responseSchema,
      },
    };

    // 常に低コスト・高速な Flash モデルを使用
    // gemini-3.8-flash を筆頭に、新しいモデルから順に自動フォールバック試行
    const defaultFlashModels = [
      'gemini-3.8-flash',
      'gemini-3.6-flash',
      'gemini-3.5-flash',
      'gemini-3.0-flash',
      'gemini-2.5-flash',
      'gemini-2.0-flash-exp',
      'gemini-1.5-flash-latest',
      'gemini-1.5-flash',
    ];

    const flashCandidates = c.env.GEMINI_MODEL
      ? Array.from(new Set([c.env.GEMINI_MODEL, ...defaultFlashModels]))
      : defaultFlashModels;

    let res: Response | null = null;
    let lastErrorText = '';

    for (const modelCandidate of flashCandidates) {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelCandidate}:generateContent?key=${c.env.GEMINI_API_KEY}`;

      let attempt = 0;
      while (attempt <= 2) {
        try {
          res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
          });

          if (res.status === 429 && attempt < 2) {
            attempt++;
            const delay = 1000 * Math.pow(2, attempt) + Math.random() * 1000;
            await new Promise((r) => setTimeout(r, delay));
            continue;
          }

          // 404 (モデル不在・廃止) の場合は次の Flash モデル候補を試す
          if (res.status === 404) {
            lastErrorText = await res.text();
            res = null;
            break;
          }

          break;
        } catch (err: any) {
          lastErrorText = err.message;
          res = null;
          break;
        }
      }

      // 成功 (200 OK) したらループを抜ける
      if (res && res.ok) {
        break;
      }
    }

    if (!res || !res.ok) {
      const errText = (await res?.text()) || lastErrorText;
      throw new Error(`API Error: ${errText}`);
    }

    const json = (await res.json()) as any;
    let aiResponseText = json.candidates[0].content.parts[0].text;
    aiResponseText = aiResponseText
      .replace(/^```(?:json)?\n?/i, '')
      .replace(/\n?```$/i, '')
      .trim();

    // データベース (スプレッドシート) へ保存
    const aiSheet = doc.sheetsByTitle[SHEET_SUMMARIES];
    if (!aiSheet) throw new Error('AI_Summaries sheet not found');

    await withRetry(() =>
      aiSheet.addRow({
        id: generateUuid(),
        session_id: sessionId,
        type: analysisType,
        summary_json: aiResponseText,
        created_at: new Date().toISOString(),
      })
    );

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
