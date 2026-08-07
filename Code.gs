const SHEET_QUESTIONS = '【意見・質問】分科会';
const SHEET_RECORDS = '【記録】分科会';
const SHEET_SYSTEM = 'System_Status';
const SHEET_SUMMARIES = 'AI_Summaries';

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
        result = updateQuestionStatus(params.rowIndex, params.status, params.isHidden);
        break;
      case 'runGeminiAnalysis':
        result = runGeminiAnalysis(params.sessionId, params.analysisType, params.scope);
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
 * ①と③を結合して質問一覧を取得（キャッシュ30秒）
 */
function getQuestions(sessionId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'questions_session_' + sessionId;
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    return JSON.parse(cachedData);
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (!qSheet) return [];
  
  const qData = qSheet.getDataRange().getValues();
  if (qData.length <= 1) return [];
  
  // システムステータスをマップ化
  const sysSheet = ss.getSheetByName(SHEET_SYSTEM);
  let sysMap = {};
  if (sysSheet && sysSheet.getLastRow() > 1) {
    const sysData = sysSheet.getDataRange().getValues();
    for (let i = 1; i < sysData.length; i++) {
      if (sysData[i][0] === SHEET_QUESTIONS) {
        sysMap[sysData[i][1]] = {
          status: sysData[i][2],
          is_hidden: sysData[i][3] === true || String(sysData[i][3]).toUpperCase() === 'TRUE'
        };
      }
    }
  }
  
  let result = [];
  for (let i = 1; i < qData.length; i++) {
    // 修正: フォーム側の分科会名（B列）からID(A1〜B8等)を抽出して比較する
    const formSessionStr = String(qData[i][1]);
    const matchSession = formSessionStr.match(/[A-Za-z]\d+/);
    const formSessionId = matchSession ? matchSession[0].toUpperCase() : formSessionStr.toUpperCase();
    
    if (formSessionId === String(sessionId).toUpperCase()) {
      const rowIndex = i + 1;
      const sysInfo = sysMap[rowIndex] || { status: '未回答', is_hidden: false };
      result.push({
        rowIndex: rowIndex,
        timestamp: qData[i][0] instanceof Date ? Utilities.formatDate(qData[i][0], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') : qData[i][0],
        raw_time: qData[i][0] instanceof Date ? qData[i][0].getTime() : new Date(qData[i][0]).getTime(), // 追加: AI分析済判定用のミリ秒データ
        name: qData[i][2],
        school: qData[i][3],
        opinion: qData[i][4],
        question: qData[i][5],
        status: sysInfo.status,
        is_hidden: sysInfo.is_hidden
      });
    }
  }
  
  try {
    cache.put(cacheKey, JSON.stringify(result), 30);
  } catch (e) {
    console.warn('Cache limit exceeded for session: ' + sessionId);
  }
  return result;
}

function updateQuestionStatus(rowIndex, status, isHidden) {
  // ★ 修正: ステータス更新処理（シートの読み書き）にも排他制御を追加
  const lock = LockService.getScriptLock();
  
  // 10秒待機（通常は数十ミリ秒で完了します）
  if (!lock.tryLock(10000)) {
    throw new Error('システム混雑のためステータスの更新に失敗しました。数秒後に再度お試しください。');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sysSheet = ss.getSheetByName(SHEET_SYSTEM);
    const sysData = sysSheet.getDataRange().getValues();
    
    let found = false;
    for (let i = 1; i < sysData.length; i++) {
      if (sysData[i][0] === SHEET_QUESTIONS && String(sysData[i][1]) === String(rowIndex)) {
        sysSheet.getRange(i + 1, 3).setValue(status);
        sysSheet.getRange(i + 1, 4).setValue(isHidden);
        sysSheet.getRange(i + 1, 5).setValue(new Date());
        found = true;
        break;
      }
    }
    
    if (!found) {
      sysSheet.appendRow([SHEET_QUESTIONS, rowIndex, status, isHidden, new Date()]);
    }
    
    // キャッシュのクリア対象を新しい分科会形式（A1〜A3, B1〜B8）に変更
    const cache = CacheService.getScriptCache();
    const sessions = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8'];
    sessions.forEach(s => cache.remove('questions_session_' + s));
    
    return true;
  } finally {
    // 成功・失敗に関わらず必ずロックを解放する
    lock.releaseLock();
  }
}

function getRecords(sessionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName(SHEET_RECORDS);
  if (!rSheet) return [];
  const rData = rSheet.getDataRange().getValues();
  let result = [];
  for (let i = 1; i < rData.length; i++) {
    // 修正: アルファベットと数字の組み合わせを抽出
    const formSessionStr = String(rData[i][1]);
    const matchSession = formSessionStr.match(/[A-Za-z]\d+/);
    const formSessionId = matchSession ? matchSession[0].toUpperCase() : formSessionStr;

    if (formSessionId === String(sessionId).toUpperCase()) {
      result.push(rData[i][2]);
    }
  }
  return result;
}

/**
 * Gemini 2.5 Flash 呼び出し処理
 */

function runGeminiAnalysis(sessionId, analysisType, scope) {
  // 関数全体のロックを削除し、APIの並列リクエストを許可
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  let prompt = '';
  let systemInstruction = '';
  let responseSchema = {};
  
  if (analysisType === 'QUESTIONS') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const qSheet = ss.getSheetByName(SHEET_QUESTIONS);
    if (!qSheet) throw new Error('質問シートがありません');
    
    let latestAITime = 0;
    if (scope === 'new') {
      const summaries = getLatestAISummaries(sessionId);
      latestAITime = summaries.qTime || 0;
    }

    const qData = qSheet.getDataRange().getValues();
    let textData = '';
    let targetCount = 0;
    
    for (let i = 1; i < qData.length; i++) {
      const formSessionStr = String(qData[i][1]);
      const matchSession = formSessionStr.match(/[A-Za-z]\d+/);
      const formSessionId = matchSession ? matchSession[0].toUpperCase() : formSessionStr;
      
      if (formSessionId === String(sessionId).toUpperCase()) {
        if (scope === 'new') {
          const raw_time = qData[i][0] instanceof Date ? qData[i][0].getTime() : new Date(qData[i][0]).getTime();
          if (raw_time <= latestAITime) continue;
        }
        textData += `質問: ${qData[i][5]} (意見: ${qData[i][4]})\n`;
        targetCount++;
      }
    }
    
    if (targetCount === 0) throw new Error(scope === 'new' ? '未分析の新しいデータがありません' : '分析するデータがありません');
    
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
    const records = getRecords(sessionId);
    if (records.length === 0) throw new Error('分析する記録データがありません');
    
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

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;
  
  const payload = {
    "system_instruction": { "parts": [{ "text": systemInstruction }] },
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": {
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
  
  // ★ 修正: スプレッドシートへの書き込み（数ミリ秒）のみを排他制御する
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

function getLatestAISummaries(sessionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aiSheet = ss.getSheetByName(SHEET_SUMMARIES);
  if (!aiSheet) return { questions: null, records: null };
  
  const data = aiSheet.getDataRange().getValues();
  let latestQuestions = null, latestRecords = null;
  let qTime = 0, rTime = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(sessionId)) {
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
    qTime: qTime // 追加: フロントエンド（latestAITime）に最新のAI実行時間を渡す
  };
}