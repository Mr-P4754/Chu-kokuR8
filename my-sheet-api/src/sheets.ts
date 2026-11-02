import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

// スプレッドシートIDを抽出するユーティリティ
export function extractSpreadsheetId(input: string): string {
  if (!input) return '';
  // URL形式の場合の抽出
  const match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : input.trim();
}

/**
 * 指数関数的バックオフ（ジッター付き）を用いて関数を実行するラッパー
 * @param operation 実行する非同期関数
 * @param maxRetries 最大再試行回数
 */
export async function withRetry<T>(operation: () => Promise<T>, maxRetries: number = 3): Promise<T> {
  let attempt = 0;
  while (attempt <= maxRetries) {
    try {
      return await operation();
    } catch (error: any) {
      const isRateLimit = error?.response?.status === 429 || error?.code === 429 || error?.message?.includes('429');
      if (isRateLimit && attempt < maxRetries) {
        attempt++;
        // 指数関数的バックオフ: 1s, 2s, 4s... + ジッター(0~1s)
        const delay = (1000 * Math.pow(2, attempt)) + (Math.random() * 1000);
        console.warn(`[Retry ${attempt}/${maxRetries}] 429 Too Many Requests. Waiting ${Math.round(delay)}ms...`);
        await new Promise(resolve => setTimeout(resolve, delay));
        continue;
      }
      throw error; // 429以外、または最大再試行回数に達した場合はエラーを投げる
    }
  }
  throw new Error('Maximum retries reached');
}

/**
 * 認証済みの GoogleSpreadsheet インスタンスを取得する
 */
export async function getAuthSpreadsheet(env: any): Promise<GoogleSpreadsheet> {
  const serviceAccountEmail = env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  // \n を実際の改行に変換
  const privateKey = env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n');
  const rawSpreadsheetId = env.SPREADSHEET_ID || 'https://docs.google.com/spreadsheets/d/166s92vo0Wk_e_PYFj9pF5oIm-0vdM5GfU6_7o5O4Aaw/edit?usp=sharing';
  const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);

  const auth = new JWT({
    email: serviceAccountEmail,
    key: privateKey,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const doc = new GoogleSpreadsheet(spreadsheetId, auth);
  
  // 初期情報の読み込みにもリトライを適用
  await withRetry(() => doc.loadInfo());
  return doc;
}
