/**
 * customer_search.js
 * 電話番号送信 → 利用履歴シート（A列〜J列）の検索結果を返信
 * - 入力パターン: 「09012345678」「09012345678 谷村」「全件 09012345678」
 * - キャッシュ: A:J列を2時間保持、SA認証で読み込み（429回避）
 */
const { google } = require('googleapis');
const { createSAAuthClient, normalizePhone, toFullWidth } = require('./lib/common');

const RIREKI_SHEET_ID   = '1A_LaiWm2QhvXk4jKINSRvKKX3-xKYLzygNnlY3ZtI9A';
const RIREKI_SHEET_NAME = '利用履歴';
const RIREKI_CACHE_TTL  = 2 * 60 * 60 * 1000; // 2時間

const rirekiCache = { allRows: null, phoneIndex: null, expiresAt: 0 };
let rirekiCacheLoading = false;

async function buildRirekiCache() {
  if (rirekiCache.allRows && Date.now() < rirekiCache.expiresAt) return;
  if (rirekiCacheLoading) {
    if (rirekiCache.allRows) return;
    await new Promise(r => setTimeout(r, 5000));
    if (rirekiCache.allRows) return;
  }
  rirekiCacheLoading = true;
  try {
    await _loadRirekiWithRetry();
  } finally {
    rirekiCacheLoading = false;
  }
}

async function _loadRirekiWithRetry(maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`[利用履歴] A:J列一括ロード開始...（${attempt}/${maxRetries}・SA認証）`);
      const auth = createSAAuthClient();
      const sheets = google.sheets({ version: 'v4', auth });
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: RIREKI_SHEET_ID,
        range: `${RIREKI_SHEET_NAME}!A:J`,
        valueRenderOption: 'FORMATTED_VALUE',
      });
      const rows = res.data.values || [];
      const phoneIndex = {};
      for (let i = 0; i < rows.length; i++) {
        const tel = normalizePhone(rows[i][9] || '');
        if (tel.length < 10) continue;
        if (!phoneIndex[tel]) phoneIndex[tel] = [];
        phoneIndex[tel].push(i);
      }
      rirekiCache.allRows = rows;
      rirekiCache.phoneIndex = phoneIndex;
      rirekiCache.expiresAt = Date.now() + RIREKI_CACHE_TTL;
      console.log(`[利用履歴] キャッシュ完了: ${rows.length}行, ${Object.keys(phoneIndex).length}件の電話番号`);
      return;
    } catch (err) {
      const is429 = err.message && err.message.includes('429');
      if (is429 && attempt < maxRetries) {
        const wait = attempt * 10 * 1000;
        console.log(`[利用履歴] 429エラー → ${wait / 1000}秒待機してリトライ`);
        await new Promise(r => setTimeout(r, wait));
        continue;
      }
      if (rirekiCache.allRows) {
        console.log(`[利用履歴] キャッシュ更新失敗（${err.message}）→ 既存データで継続`);
        rirekiCache.expiresAt = Date.now() + 10 * 60 * 1000;
        return;
      }
      throw err;
    }
  }
}

function invalidateCache() {
  rirekiCache.expiresAt = 0;
}

function preloadPhoneIndex() {
  setImmediate(() => buildRirekiCache().catch(e => console.error('[利用履歴] プリロードエラー:', e.message)));
}

// 電話番号で利用履歴を検索（J列に部分一致＝シート側CONTAINSと同じ挙動）
async function searchRirekiByPhone(phone, allRecords = false, nameKeyword = '') {
  const normalized = normalizePhone(phone);
  if (normalized.length < 9) return null;
  await buildRirekiCache();
  const rows = rirekiCache.allRows || [];
  const kw = nameKeyword ? String(nameKeyword).trim() : '';
  const targets = [];
  for (let i = rows.length - 1; i >= 0; i--) {
    const tel = normalizePhone((rows[i] || [])[9] || '');
    if (!tel || !tel.includes(normalized)) continue;
    if (kw) {
      const name = String((rows[i] || [])[2] || '');
      if (!name.includes(kw)) continue;
    }
    targets.push(i);
  }
  if (targets.length === 0) return { total: 0, rows: [], filtered: !!kw };
  const dataRows = targets.map(i => (rows[i] || []).slice(0, 7));
  return { total: targets.length, rows: dataRows, filtered: !!kw };
}

function formatRirekiResult(phone, result, allRecords = false, nameKeyword = '') {
  const { total, rows } = result;
  const label = `全${rows.length}件`;
  const headerKey = nameKeyword ? `${phone} / ${nameKeyword}` : phone;
  const header = `📋 利用履歴: ${headerKey}\n合計 ${total} 件（${label}）`;
  const lines = rows.map((r, i) => {
    const date  = r[0] || '';
    const store = r[1] || '';
    const name  = r[2] || '';
    const type  = r[3] || '';
    const cours = toFullWidth(r[4] || '');
    const op    = toFullWidth(r[5] || '');
    const memo  = r[6] || '';
    const line1 = [date, store].filter(v => v).join(' ');
    const line2 = [name, type, cours, op, memo].filter(v => v).join(' ');
    return `${i + 1}. ${line1}\n   ${line2}`;
  });
  return [header, ...lines].join('\n');
}

// 電話番号パターン: 「09012345678」「09012345678 谷村」「全件 09012345678」
function parsePhoneCommand(text) {
  const normalized = text.replace(/[　]/g, ' ').trim();
  const tokens = normalized.split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return null;

  if (tokens[0] === '全件' && tokens.length === 2) {
    const phone = tokens[1];
    const d = normalizePhone(phone);
    if (/^[0-9０-９\-ー－]+$/.test(phone) && d.length >= 9 && d.length <= 11 && d.startsWith('0')) {
      return { phone, allRecords: true, nameKeyword: '' };
    }
  }

  const head = tokens[0];
  const headDigits = normalizePhone(head);
  if (/^[0-9０-９\-ー－]+$/.test(head) && headDigits.length >= 9 && headDigits.length <= 11 && headDigits.startsWith('0')) {
    if (tokens.length === 1) return { phone: head, allRecords: true, nameKeyword: '' };
    if (tokens.length === 2) {
      if (tokens[1] === '全件') return { phone: head, allRecords: true, nameKeyword: '' };
      return { phone: head, allRecords: true, nameKeyword: tokens[1] };
    }
  }
  return null;
}

// LINEイベントを処理（マッチしなければ false を返す）
async function handleEvent(event, text, client) {
  const phoneCmd = parsePhoneCommand(text);
  if (!phoneCmd) return false;

  const replyToken = event.replyToken;
  const isGroup = event.source?.type === 'group';
  console.log(`[LINE] 電話番号検索${isGroup ? '(グループ)' : ''}: ${phoneCmd.phone} 全件=${phoneCmd.allRecords} 名前=${phoneCmd.nameKeyword || '(なし)'}`);
  setImmediate(async () => {
    try {
      const result = await searchRirekiByPhone(phoneCmd.phone, phoneCmd.allRecords, phoneCmd.nameKeyword);
      console.log(`[利用履歴] 検索結果: ${result ? result.total + '件' : '該当なし'}`);
      const headerKey = phoneCmd.nameKeyword ? `${phoneCmd.phone} / ${phoneCmd.nameKeyword}` : phoneCmd.phone;
      const msg = (result && result.total > 0)
        ? formatRirekiResult(phoneCmd.phone, result, phoneCmd.allRecords, phoneCmd.nameKeyword)
        : `📋 利用履歴: ${headerKey}\n該当なし`;
      await client.replyMessage({ replyToken, messages: [{ type: 'text', text: msg }] });
      console.log(`[利用履歴] LINE送信完了`);
    } catch (err) {
      console.error('[利用履歴] エラー詳細:', err.code || err.statusCode || '', err.message);
    }
  });
  return true;
}

module.exports = {
  handleEvent,
  searchRirekiByPhone,
  formatRirekiResult,
  parsePhoneCommand,
  preloadPhoneIndex,
  invalidateCache,
};
