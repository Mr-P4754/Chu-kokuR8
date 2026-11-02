// 分科会IDと【意見・質問】シート名の対応マップ
const SESSION_SHEET_MAP = {
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
  'B-4②': '【意見・質問】B-4分科会②'
};

const SHEET_RECORDS = '【記録】分科会';
const SHEET_SYSTEM = 'System_Status';
const SHEET_SUMMARIES = 'AI_Summaries';

/**
 * 文字列から正規化された分科会ID（例: 'A-1', 'B-1①'）を抽出する
 */
function normalizeSessionId(rawText) {
  if (!rawText) return '';
  const trimmed = String(rawText).trim();
  if (SESSION_SHEET_MAP[trimmed]) return trimmed;

  const normalized = trimmed
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[-‐‑‒–—―ー－−]/g, '-')
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
function getQuestionSheetTitle(sessionId) {
  const normalizedId = normalizeSessionId(sessionId);
  return SESSION_SHEET_MAP[normalizedId] || `【意見・質問】${normalizedId}分科会`;
}

/**
 * 初期アクセス時にシートを検証し、不足・破損があれば自己修復する
 */
function verifyAndHealSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // System_Statusシートの検証と修復
  let sysSheet = ss.getSheetByName(SHEET_SYSTEM);
  const sysHeader = ['sheet_name', 'row_index', 'status', 'is_hidden', 'updated_at'];
  if (!sysSheet) {
    sysSheet = ss.insertSheet(SHEET_SYSTEM);
    sysSheet.appendRow(sysHeader);
  } else {
    if (sysSheet.getLastRow() === 0 || sysSheet.getLastColumn() < 5) {
      sysSheet.getRange(1, 1, 1, 5).setValues([sysHeader]);
    } else {
      const header = sysSheet.getRange(1, 1, 1, 5).getValues()[0];
      if (JSON.stringify(header) !== JSON.stringify(sysHeader)) {
        sysSheet.getRange(1, 1, 1, 5).setValues([sysHeader]);
      }
    }
  }
  
  // AI_Summariesシートの検証と修復
  let aiSheet = ss.getSheetByName(SHEET_SUMMARIES);
  const aiHeader = ['id', 'session_id', 'type', 'summary_json', 'created_at'];
  if (!aiSheet) {
    aiSheet = ss.insertSheet(SHEET_SUMMARIES);
    aiSheet.appendRow(aiHeader);
  } else {
    if (aiSheet.getLastRow() === 0 || aiSheet.getLastColumn() < 5) {
      aiSheet.getRange(1, 1, 1, 5).setValues([aiHeader]);
    } else {
      const header = aiSheet.getRange(1, 1, 1, 5).getValues()[0];
      if (JSON.stringify(header) !== JSON.stringify(aiHeader)) {
        aiSheet.getRange(1, 1, 1, 5).setValues([aiHeader]);
      }
    }
  }
}

// APIエンドポイント（GETリクエスト用）
function doGet(e) {
  return handleRequest(e, 'GET');
}

// APIエンドポイント（POSTリクエスト用）
function doPost(e) {
  return handleRequest(e, 'POST');
}

// リクエストの振り分けと実行
function handleRequest(e, method) {
  verifyAndHealSheets();

  let action = e.parameter.action;
  let params = e.parameter;

  // POSTの場合はJSONペイロードをパース
  if (method === 'POST' && e.postData && e.postData.contents) {
    try {
      const payload = JSON.parse(e.postData.contents);
      action = action || payload.action;
      params = { ...params, ...payload };
    } catch (err) {
      return createJsonResponse({ success: false, error: 'JSONの解析に失敗しました' });
    }
  }

  if (!action) {
    return createJsonResponse({ success: false, error: 'アクションが指定されていません' });
  }

  try {
    let result = null;
    switch (action) {
      case 'getConfig':
        // 初期設定情報の取得（フォームURLなど）
        result = { formUrl: PropertiesService.getScriptProperties().getProperty('FORM_URL') || '' };
        break;
      case 'getQuestions':
        result = getQuestions(params.sessionId);
        break;
      case 'verifyAdmin':
        result = verifyAdmin(params.pin);
        break;
      case 'updateQuestionStatus':
        result = updateQuestionStatus(params.sessionId, params.rowIndex, params.status, params.isHidden);
        break;
      case 'runGeminiAnalysis':
        result = runGeminiAnalysis(params.sessionId, params.analysisType, params.scope);
        break;
      case 'deleteRecords':
        result = deleteRecords(params.sessionId);
        break;
      case 'getLatestAISummaries':
        result = getLatestAISummaries(params.sessionId);
        break;
      default:
        return createJsonResponse({ success: false, error: '無効なアクションです: ' + action });
    }
    return createJsonResponse({ success: true, data: result });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.message });
  }
}

// JSONレスポンスの生成
function createJsonResponse(responseObject) {
  // CORS対策として、GASはContentServiceに対して自動的に適切なヘッダーを付与します
  return ContentService.createTextOutput(JSON.stringify(responseObject))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyAdmin(pin) {
  const correctPin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || '1204';
  return pin === correctPin;
}

/**
 * 分科会シートとSystem_Statusシートを結合して質問一覧を取得（キャッシュ30秒）
 */
function getQuestions(rawSessionId) {
  const sessionId = normalizeSessionId(rawSessionId);
  const questionSheetTitle = getQuestionSheetTitle(sessionId);

  const cache = CacheService.getScriptCache();
  const cacheKey = 'questions_session_' + sessionId;
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    return JSON.parse(cachedData);
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(questionSheetTitle);
  if (!qSheet) return [];
  
  const qData = qSheet.getDataRange().getValues();
  if (qData.length <= 1) return [];
  
  // システムステータスをマップ化
  const sysSheet = ss.getSheetByName(SHEET_SYSTEM);
  let sysMap = {};
  if (sysSheet && sysSheet.getLastRow() > 1) {
    const sysData = sysSheet.getDataRange().getValues();
    for (let i = 1; i < sysData.length; i++) {
      if (sysData[i][0] === questionSheetTitle) {
        sysMap[sysData[i][1]] = {
          status: sysData[i][2],
          is_hidden: sysData[i][3] === true || String(sysData[i][3]).toUpperCase() === 'TRUE'
        };
      }
    }
  }
  
  const headers = qData[0].map(h => String(h).trim());
  const hasSessionColumn = headers.some(h => h.includes('分科会')) || headers.length >= 6;

  let result = [];
  for (let i = 1; i < qData.length; i++) {
    const row = qData[i];
    const timestampVal = row[0];
    const name = hasSessionColumn ? row[2] : row[1];
    const school = hasSessionColumn ? row[3] : row[2];
    const opinion = hasSessionColumn ? row[4] : row[3];
    const question = hasSessionColumn ? row[5] : row[4];

    if (!timestampVal && !name && !opinion && !question) continue;

    const rowIndex = i + 1;
    const sysInfo = sysMap[rowIndex] || { status: '未回答', is_hidden: false };
    result.push({
      rowIndex: rowIndex,
      timestamp: timestampVal instanceof Date ? Utilities.formatDate(timestampVal, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') : timestampVal,
      raw_time: timestampVal instanceof Date ? timestampVal.getTime() : new Date(timestampVal).getTime(),
      name: name,
      school: school,
      opinion: opinion,
      question: question,
      status: sysInfo.status,
      is_hidden: sysInfo.is_hidden
    });
  }
  
  try {
    cache.put(cacheKey, JSON.stringify(result), 30);
  } catch (e) {
    console.warn('Cache limit exceeded for session: ' + sessionId);
  }
  return result;
}

/**
 * 質問の回答ステータス・非表示フラグを更新する
 */
function updateQuestionStatus(rawSessionId, rowIndex, status, isHidden) {
  const lock = LockService.getScriptLock();
  
  if (!lock.tryLock(10000)) {
    throw new Error('システム混雑のためステータスの更新に失敗しました。数秒後に再度お試しください。');
  }

  try {
    const targetSheetTitle = rawSessionId ? getQuestionSheetTitle(rawSessionId) : '【意見・質問】A-1分科会';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sysSheet = ss.getSheetByName(SHEET_SYSTEM);
    const sysData = sysSheet.getDataRange().getValues();
    
    let found = false;
    for (let i = 1; i < sysData.length; i++) {
      if (sysData[i][0] === targetSheetTitle && String(sysData[i][1]) === String(rowIndex)) {
        sysSheet.getRange(i + 1, 3).setValue(status);
        sysSheet.getRange(i + 1, 4).setValue(isHidden);
        sysSheet.getRange(i + 1, 5).setValue(new Date());
        found = true;
        break;
      }
    }
    
    if (!found) {
      sysSheet.appendRow([targetSheetTitle, rowIndex, status, isHidden, new Date()]);
    }
    
    // キャッシュのクリア対象を新しい分科会形式に変更
    const cache = CacheService.getScriptCache();
    const sessions = Object.keys(SESSION_SHEET_MAP);
    sessions.forEach(s => cache.remove('questions_session_' + s));
    
    return true;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 記録シートから該当分科会の文字起こしデータを取得する
 */
function getRecords(rawSessionId) {
  const sessionId = normalizeSessionId(rawSessionId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName(SHEET_RECORDS);
  if (!rSheet) return [];
  const rData = rSheet.getDataRange().getValues();
  let result = [];
  for (let i = 1; i < rData.length; i++) {
    const formSessionStr = String(rData[i][1]);
    const formSessionId = normalizeSessionId(formSessionStr);
    const recordText = String(rData[i][2] || '').trim();

    if (formSessionId === sessionId && recordText) {
      result.push(recordText);
    }
  }
  return result;
}

/**
 * 該当分科会の記録データおよび記録要約データをスプレッドシートから削除する
 */
function deleteRecords(rawSessionId) {
  const sessionId = normalizeSessionId(rawSessionId);
  if (!sessionId) throw new Error('sessionId が指定されていません');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('システム混雑のため削除に失敗しました。数秒後に再度お試しください。');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let deletedRecordsCount = 0;
    let deletedSummariesCount = 0;

    // 1. 【記録】分科会シートから該当分科会の行を末尾から逆順に削除
    const rSheet = ss.getSheetByName(SHEET_RECORDS);
    if (rSheet && rSheet.getLastRow() > 1) {
      const rData = rSheet.getDataRange().getValues();
      for (let i = rData.length - 1; i >= 1; i--) {
        const formSessionId = normalizeSessionId(String(rData[i][1]));
        if (formSessionId === sessionId) {
          rSheet.deleteRow(i + 1);
          deletedRecordsCount++;
        }
      }
    }

    // 2. AI_Summariesシートから該当分科会のRECORDS要約行を末尾から逆順に削除
    const aiSheet = ss.getSheetByName(SHEET_SUMMARIES);
    if (aiSheet && aiSheet.getLastRow() > 1) {
      const aiData = aiSheet.getDataRange().getValues();
      for (let i = aiData.length - 1; i >= 1; i--) {
        const rowSessionId = normalizeSessionId(String(aiData[i][1]));
        const rowType = String(aiData[i][2]);
        if (rowSessionId === sessionId && rowType === 'RECORDS') {
          aiSheet.deleteRow(i + 1);
          deletedSummariesCount++;
        }
      }
    }

    return {
      deletedRecordsCount: deletedRecordsCount,
      deletedSummariesCount: deletedSummariesCount
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Gemini 3.8 Flash 呼び出し処理
 */
function runGeminiAnalysis(rawSessionId, analysisType, scope) {
  const sessionId = normalizeSessionId(rawSessionId);
  const questionSheetTitle = getQuestionSheetTitle(sessionId);

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  let prompt = '';
  let systemInstruction = '';
  let responseSchema = {};
  
  if (analysisType === 'QUESTIONS') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const qSheet = ss.getSheetByName(questionSheetTitle);
    if (!qSheet) throw new Error(`質問シート（${questionSheetTitle}）がありません`);
    
    let latestAITime = 0;
    if (scope === 'new') {
      const summaries = getLatestAISummaries(sessionId);
      latestAITime = summaries.qTime || 0;
    }

    const qData = qSheet.getDataRange().getValues();
    if (qData.length <= 1) throw new Error('分析するデータがありません');

    const headers = qData[0].map(h => String(h).trim());
    const hasSessionColumn = headers.some(h => h.includes('分科会')) || headers.length >= 6;

    let textData = '';
    let targetCount = 0;
    
    for (let i = 1; i < qData.length; i++) {
      const row = qData[i];
      const opinion = hasSessionColumn ? row[4] : row[3];
      const question = hasSessionColumn ? row[5] : row[4];
      if (!opinion && !question) continue;

      if (scope === 'new') {
        const raw_time = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();
        if (raw_time <= latestAITime) continue;
      }
      textData += `質問: ${question} (意見: ${opinion})\n`;
      targetCount++;
    }
    
    if (targetCount === 0) throw new Error(scope === 'new' ? '未分析の新しいデータがありません' : '分析するデータがありません');
    
    systemInstruction = [
      'あなたは教育研究大会の質疑応答モデレーターです。',
      '提供された質問と意見のデータのみに基づき、論点を整理してスキーマに従って出力してください。',
      '【厳守事項】提供されたテキストに含まれていない内容や一般的な教育論を勝手に推測・創作（ハルシネーション）しないでください。'
    ].join('\n');
    prompt = `以下の分科会${sessionId}の質問・意見データのみを分析してください（データにない内容は一切追加しないこと）:\n---\n${textData}---`;
    responseSchema = {
      "type": "OBJECT",
      "properties": {
        "overall_summary": { "type": "STRING", "description": "提供されたデータに基づく全体傾向(最大200字程度)" },
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
    const records = getRecords(sessionId);
    if (records.length === 0) throw new Error('分析する記録データがありません');
    
    systemInstruction = [
      'あなたは教育研究大会の正確な記録係です。以下の【絶対厳守ルール】に従って、提供された「分科会の音声文字起こしデータ」から要点を整理し、スキーマに従って出力してください。',
      '【絶対厳守ルール】',
      '1. 提供された文字起こしテキストに明記されている事実・発言内容のみに基づいて要約を作成してください。',
      '2. 文字起こしテキストに存在しない内容（一般的な教育用語、ICT活用、個別最適な学び、架空の発表・質疑応答など）を推測・補完・創作（ハルシネーション）することは絶対に禁止します。',
      '3. 入力が「テスト」などの単語のみ、または実質的な発表・議論の内容が読み取れない短いテキストである場合は、絶対に架空の文章を作らず、session_summary には「有効な文字起こしデータが不足しているため要約を作成できません（入力内容：…）。」のように入力された事実のみをそのまま記載し、key_takeaways は空配列 []、discussion_points は「該当なし」としてください。'
    ].join('\n');
    prompt = `以下は分科会${sessionId}の実際の文字起こしデータです。このテキストに書かれている内容のみを忠実に要約してください（書かれていない事柄は一切追加・創作しないでください）:\n---\n${records.join('\n')}\n---`;
    responseSchema = {
      "type": "OBJECT",
      "properties": {
        "session_summary": { "type": "STRING", "description": "提供された文字起こしに基づく分科会全体の要約(最大400字程度。内容が短い・テスト等の場合は無理に膨らませず入力された事実のみを短く記載すること)" },
        "key_takeaways": { "type": "ARRAY", "items": { "type": "STRING" }, "description": "文字起こしから読み取れる重要なポイント(該当する内容がない場合は空配列)" },
        "discussion_points": { "type": "STRING", "description": "文字起こし内で実際に議論されたことや課題(該当する内容がない場合は「該当なし」)" }
      },
      "required": ["session_summary", "key_takeaways", "discussion_points"]
    };
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${apiKey}`;
  
  const payload = {
    "system_instruction": { "parts": [{ "text": systemInstruction }] },
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": {
      "temperature": 0.1,
      "response_mime_type": "application/json",
      "response_schema": responseSchema
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  // 429エラー対策のジッター制御付き再試行ロジック
  let response;
  let retries = 0;
  const maxRetries = 3;
  
  while (retries <= maxRetries) {
    response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 429 && retries < maxRetries) {
      retries++;
      Utilities.sleep((1000 * Math.pow(2, retries)) + (Math.random() * 1000));
      continue;
    }
    break;
  }
  
  if (response.getResponseCode() !== 200) {
    let errorDetail = response.getContentText();
    try {
      const errJson = JSON.parse(errorDetail);
      if (errJson.error && errJson.error.message) {
        errorDetail = `${errJson.error.code} - ${errJson.error.message}`;
      }
    } catch (e) {}
    throw new Error('API Error: ' + errorDetail);
  }

  const json = JSON.parse(response.getContentText());
  let aiResponseText = json.candidates[0].content.parts[0].text;
  
  aiResponseText = aiResponseText.replace(/^```(?:json)?\n?/i, '').replace(/\n?```$/i, '').trim();
  
  // スプレッドシートへの書き込みのみを排他制御する
  const writeLock = LockService.getScriptLock();
  if (writeLock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const aiSheet = ss.getSheetByName(SHEET_SUMMARIES);
      aiSheet.appendRow([Utilities.getUuid(), sessionId, analysisType, aiResponseText, new Date()]);
    } finally {
      writeLock.releaseLock();
    }
  } else {
    throw new Error('AIの解析は完了しましたが、システム混雑により結果の保存に失敗しました。もう一度実行してください。');
  }
  
  return aiResponseText;
}

function getLatestAISummaries(rawSessionId) {
  const sessionId = normalizeSessionId(rawSessionId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aiSheet = ss.getSheetByName(SHEET_SUMMARIES);
  if (!aiSheet) return { questions: null, records: null };
  
  const data = aiSheet.getDataRange().getValues();
  let latestQuestions = null, latestRecords = null;
  let qTime = 0, rTime = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (normalizeSessionId(String(data[i][1])) === sessionId) {
      const type = data[i][2];
      const time = new Date(data[i][4]).getTime();
      if (type === 'QUESTIONS' && time > qTime) {
        latestQuestions = data[i][3];
        qTime = time;
      } else if (type === 'RECORDS' && time > rTime) {
        latestRecords = data[i][3];
        rTime = time;
      }
    }
  }
  
  // 安全なパース処理（フェイルセーフ）
  const safeParse = (str) => {
    try {
      return str ? JSON.parse(str) : null;
    } catch (e) {
      console.warn('JSON Parse Error:', e);
      return null;
    }
  };
  
  return {
    questions: safeParse(latestQuestions),
    records: safeParse(latestRecords),
    qTime: qTime
  };
}
