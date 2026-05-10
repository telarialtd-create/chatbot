/**
 * lib/common.js
 * 全機能で共有する utility（認証・日付・PNG下地・スプシ検索・LINEクライアント等）
 */
const { google } = require('googleapis');
const { messagingApi } = require('@line/bot-sdk');
const { registerFont } = require('canvas');
const fs = require('fs');
const path = require('path');

// ── 定数 ─────────────────────────────────────────────────
const NIPPO_FOLDER_ID = '16R1BK5NnvYkH4Eqh6t51OGQl0tXVJ3If';
const TEMP_DIR = path.join(__dirname, '..', 'public', 'temp');
const FONT_PATH = path.join(__dirname, '..', 'fonts', 'NotoSansJP.ttf');

// 日本語フォント登録（重複登録は canvas 側で無害）
try { registerFont(FONT_PATH, { family: 'NotoJP' }); } catch (_) {}

// ── LINE クライアント ─────────────────────────────────────
function createLineClient() {
  return new messagingApi.MessagingApiClient({
    channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  });
}

// ── Google 認証（OAuth） ──────────────────────────────────
function createAuthClient() {
  let client_id, client_secret, refresh_token;

  const credsPath = path.join(process.env.HOME, '.config/gdrive-server-credentials.json');
  const keysPath  = path.join(process.env.HOME, '.config/gcp-oauth.keys.json');

  if (fs.existsSync(credsPath) && fs.existsSync(keysPath)) {
    const oauthKeys   = JSON.parse(fs.readFileSync(keysPath));
    const credentials = JSON.parse(fs.readFileSync(credsPath));
    const installed   = oauthKeys.installed || oauthKeys.web;
    client_id     = installed.client_id;
    client_secret = installed.client_secret;
    refresh_token = credentials.refresh_token;
  } else {
    client_id     = process.env.GOOGLE_CLIENT_ID;
    client_secret = process.env.GOOGLE_CLIENT_SECRET;
    refresh_token = process.env.GOOGLE_REFRESH_TOKEN;
  }

  const client = new google.auth.OAuth2(client_id, client_secret, 'http://localhost');
  client.setCredentials({ refresh_token });
  return client;
}

// ── Service Account 認証（利用履歴の読み取り専用、429回避用） ──
function createSAAuthClient() {
  const saKeyPath = path.join(process.env.HOME, '.config/chatbot-service-account.json');
  if (fs.existsSync(saKeyPath)) {
    const key = JSON.parse(fs.readFileSync(saKeyPath));
    return new google.auth.GoogleAuth({
      credentials: key,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
  }
  return createAuthClient();
}

// ── 日付ユーティリティ ────────────────────────────────────
function getJSTDate() {
  const now = new Date();
  return new Date(now.getTime() + 9 * 60 * 60 * 1000);
}

// 朝3時〜翌2時59分 = 当日（3時以前は前日扱い）※JST基準
function getSheetBusinessDate() {
  const jst = getJSTDate();
  if (jst.getUTCHours() < 3) {
    const d = new Date(jst);
    d.setUTCDate(d.getUTCDate() - 1);
    return d;
  }
  return jst;
}

function dateToNippoName(date) {
  return `${date.getUTCFullYear()}年${date.getUTCMonth() + 1}月${date.getUTCDate()}日`;
}

// 「4月7日」→「2026年4月7日」
function normalizeDateStr(dateStr) {
  if (/\d{4}年/.test(dateStr)) return dateStr;
  const jst = getJSTDate();
  const year = jst.getUTCFullYear();
  return `${year}年${dateStr}`;
}

// 「4/27」「2026/4/27」「4-27」→「4月27日」
function normalizeSlashDate(raw) {
  if (!raw) return raw;
  const s = String(raw).trim();
  if (/月.*日/.test(s)) return s;
  let m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if (m) return `${m[1]}年${parseInt(m[2],10)}月${parseInt(m[3],10)}日`;
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})$/);
  if (m) return `${parseInt(m[1],10)}月${parseInt(m[2],10)}日`;
  return s;
}

// ── 文字列正規化 ──────────────────────────────────────────
function normalizePhone(tel) {
  if (!tel) return '';
  const s = tel.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  return s.replace(/[^0-9]/g, '');
}

function parseAmount(raw) {
  if (raw === null || raw === undefined) return null;
  const cleaned = String(raw).replace(/[,，\s　円¥￥]/g, '');
  if (!cleaned) return null;
  const n = parseInt(cleaned, 10);
  return Number.isNaN(n) ? null : n;
}

function removeRoomNumber(place) {
  return place.replace(/[\s　]+[\d０-９]{3,4}$/, '').trim();
}

// 半角カタカナ→全角カタカナ
function toFullWidth(str) {
  const map = {
    'ｦ':'ヲ','ｧ':'ァ','ｨ':'ィ','ｩ':'ゥ','ｪ':'ェ','ｫ':'ォ','ｬ':'ャ','ｭ':'ュ','ｮ':'ョ','ｯ':'ッ',
    'ｰ':'ー','ｱ':'ア','ｲ':'イ','ｳ':'ウ','ｴ':'エ','ｵ':'オ','ｶ':'カ','ｷ':'キ','ｸ':'ク','ｹ':'ケ',
    'ｺ':'コ','ｻ':'サ','ｼ':'シ','ｽ':'ス','ｾ':'セ','ｿ':'ソ','ﾀ':'タ','ﾁ':'チ','ﾂ':'ツ','ﾃ':'テ',
    'ﾄ':'ト','ﾅ':'ナ','ﾆ':'ニ','ﾇ':'ヌ','ﾈ':'ネ','ﾉ':'ノ','ﾊ':'ハ','ﾋ':'ヒ','ﾌ':'フ','ﾍ':'ヘ',
    'ﾎ':'ホ','ﾏ':'マ','ﾐ':'ミ','ﾑ':'ム','ﾒ':'メ','ﾓ':'モ','ﾔ':'ヤ','ﾕ':'ユ','ﾖ':'ヨ','ﾗ':'ラ',
    'ﾘ':'リ','ﾙ':'ル','ﾚ':'レ','ﾛ':'ロ','ﾜ':'ワ','ｦ':'ヲ','ﾝ':'ン',
  };
  const dakuMap = {
    'カ':'ガ','キ':'ギ','ク':'グ','ケ':'ゲ','コ':'ゴ','サ':'ザ','シ':'ジ','ス':'ズ','セ':'ゼ','ソ':'ゾ',
    'タ':'ダ','チ':'ヂ','ツ':'ヅ','テ':'デ','ト':'ド','ハ':'バ','ヒ':'ビ','フ':'ブ','ヘ':'ベ','ホ':'ボ',
    'ウ':'ヴ',
  };
  const handakuMap = {
    'ハ':'パ','ヒ':'ピ','フ':'プ','ヘ':'ペ','ホ':'ポ',
  };
  let result = '';
  for (let i = 0; i < str.length; i++) {
    const ch   = map[str[i]] || str[i];
    const next = str[i + 1];
    if (next === 'ﾞ' && dakuMap[ch])         { result += dakuMap[ch];    i++; }
    else if (next === 'ﾟ' && handakuMap[ch]) { result += handakuMap[ch]; i++; }
    else                                     { result += ch; }
  }
  return result;
}

function toRgba(c, fallback = '#ffffff') {
  if (!c) return fallback;
  const r = Math.round((c.red   || 0) * 255);
  const g = Math.round((c.green || 0) * 255);
  const b = Math.round((c.blue  || 0) * 255);
  return `rgb(${r},${g},${b})`;
}

// ── スプシ検索 ────────────────────────────────────────────
async function findSpreadsheetId(date) {
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const dateStr = dateToNippoName(date);
  const res = await drive.files.list({
    q: `'${NIPPO_FOLDER_ID}' in parents and name contains '${dateStr}'`,
    fields: 'files(id, name)',
    orderBy: 'createdTime desc',
    pageSize: 5,
  });
  const files = res.data.files || [];
  console.log(`[共通] 日報検索 (${files.length}件):`, files.map(f => f.name).join(', '));
  const file = files[0];
  if (!file) throw new Error(`本日のファイルが見つかりません: ${dateStr}`);
  return file.id;
}

// 日付文字列でスプシ検索（明細・顧客更新で使用、テストモードあり）
async function findSpreadsheetByDateStr(dateStr) {
  if (process.env.MEISAI_TEST_ID) {
    console.log(`[共通] テストモード MEISAI_TEST_ID=${process.env.MEISAI_TEST_ID}`);
    return process.env.MEISAI_TEST_ID;
  }
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const normalizedDate = normalizeDateStr(dateStr);
  console.log(`[共通] 日付検索: "${dateStr}" → "${normalizedDate}"`);
  const res = await drive.files.list({
    q: `'${NIPPO_FOLDER_ID}' in parents and name contains '${normalizedDate}'`,
    fields: 'files(id, name)',
    orderBy: 'createdTime desc',
    pageSize: 10,
  });
  const files = res.data.files || [];
  if (!files.length) throw new Error(`ファイルが見つかりません: ${normalizedDate}`);
  files.sort((a, b) => a.name.length - b.name.length);
  const file = files[0];
  console.log(`[共通] 選択: ${file.name} (${file.id})`);
  return file.id;
}

// LINE Push送信先（個人優先、なければENV）
function getPushTarget(event) {
  if (event.source?.userId) return event.source.userId;
  return process.env.LINE_USER_ID;
}

module.exports = {
  // 定数
  NIPPO_FOLDER_ID, TEMP_DIR, FONT_PATH,
  // 認証
  createAuthClient, createSAAuthClient, createLineClient,
  // 日付
  getJSTDate, getSheetBusinessDate, dateToNippoName, normalizeDateStr, normalizeSlashDate,
  // 文字列
  normalizePhone, parseAmount, removeRoomNumber, toFullWidth, toRgba,
  // スプシ検索
  findSpreadsheetId, findSpreadsheetByDateStr,
  // LINE
  getPushTarget,
};
