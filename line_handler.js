/**
 * line_handler.js
 * LINEで「(教えて)」と送ると、当日の日報スプレッドシートCA3:CR32のスクショを返す
 *
 * 日付ルール: 朝6時〜翌5時59分 = 同日（29時制）
 */
require('dotenv').config();
const { middleware, messagingApi } = require('@line/bot-sdk');
const { google } = require('googleapis');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

const NIPPO_FOLDER_ID = '1isPYyiUqyWXnS1mtpE1_YWJ9QZBTemdJ';
const TARGET_GID = 1873674341;
const RANGE = 'CA3:CR32';
const TEMP_DIR = path.join(__dirname, 'public', 'temp');

// ── LINE クライアント ─────────────────────────────────────
const lineConfig = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN || '',
  channelSecret: process.env.LINE_CHANNEL_SECRET || '',
};

function createLineClient() {
  return new messagingApi.MessagingApiClient({
    channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  });
}

// ── Google 認証 ───────────────────────────────────────────
function createAuthClient() {
  let client_id, client_secret, refresh_token, access_token;

  if (process.env.GOOGLE_CLIENT_ID) {
    client_id     = process.env.GOOGLE_CLIENT_ID;
    client_secret = process.env.GOOGLE_CLIENT_SECRET;
    refresh_token = process.env.GOOGLE_REFRESH_TOKEN;
    access_token  = process.env.GOOGLE_ACCESS_TOKEN || null;
  } else {
    const oauthKeys   = JSON.parse(fs.readFileSync(path.join(process.env.HOME, '.config/gcp-oauth.keys.json')));
    const credentials = JSON.parse(fs.readFileSync(path.join(process.env.HOME, '.config/gdrive-server-credentials.json')));
    client_id     = oauthKeys.installed.client_id;
    client_secret = oauthKeys.installed.client_secret;
    refresh_token = credentials.refresh_token;
    access_token  = credentials.access_token;
  }

  const client = new google.auth.OAuth2(client_id, client_secret, 'http://localhost');
  client.setCredentials({ access_token, refresh_token });
  return client;
}

// ── 日付ユーティリティ ────────────────────────────────────
// 朝6時〜翌5時59分 = 当日（6時以前は前日扱い）
function getSheetBusinessDate() {
  const now = new Date();
  if (now.getHours() < 6) {
    const d = new Date(now);
    d.setDate(d.getDate() - 1);
    return d;
  }
  return now;
}

function dateToNippoName(date) {
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
}

// ── Drive からスプレッドシート ID を取得 ──────────────────
async function findSpreadsheetId(date) {
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const dateStr = dateToNippoName(date);

  const res = await drive.files.list({
    q: `'${NIPPO_FOLDER_ID}' in parents and name contains '${dateStr}'`,
    fields: 'files(id, name)',
    pageSize: 5,
  });

  const file = res.data.files?.[0];
  if (!file) throw new Error(`本日のファイルが見つかりません: ${dateStr}`);
  console.log(`[LINE] ファイル発見: ${file.name} (${file.id})`);
  return file.id;
}

// ── Sheets API でセルデータ取得（値＋書式） ───────────────
async function fetchCellsData(spreadsheetId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // シート名を GID から解決
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = meta.data.sheets.find(s => s.properties.sheetId === TARGET_GID);
  if (!sheet) throw new Error(`シートが見つかりません (gid=${TARGET_GID})`);
  const sheetName = sheet.properties.title;

  // 値＋書式を一括取得
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [`'${sheetName}'!${RANGE}`],
    includeGridData: true,
  });

  return res.data.sheets?.[0]?.data?.[0];
}

// ── セルデータ → HTML テーブル ────────────────────────────
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function colorStyle(colorObj) {
  if (!colorObj) return '';
  const r = Math.round((colorObj.red   || 0) * 255);
  const g = Math.round((colorObj.green || 0) * 255);
  const b = Math.round((colorObj.blue  || 0) * 255);
  return `rgb(${r},${g},${b})`;
}

function buildHtml(gridData) {
  const rows = gridData?.rowData || [];

  const tableRows = rows.map(row => {
    const cells = row.values || [];
    const tds = cells.map(cell => {
      const value = cell.formattedValue ?? '';
      const fmt   = cell.effectiveFormat || {};
      const tf    = fmt.textFormat || {};

      const styles = [];

      // 背景色（白でなければ適用）
      const bg = fmt.backgroundColor;
      if (bg && !(bg.red === 1 && bg.green === 1 && bg.blue === 1)) {
        styles.push(`background-color:${colorStyle(bg)}`);
      }

      // 文字色
      const fg = tf.foregroundColor;
      if (fg) styles.push(`color:${colorStyle(fg)}`);

      // 太字・フォントサイズ
      if (tf.bold)     styles.push('font-weight:bold');
      if (tf.fontSize) styles.push(`font-size:${tf.fontSize}pt`);

      // 水平配置
      const ha = fmt.horizontalAlignment;
      if (ha === 'CENTER') styles.push('text-align:center');
      else if (ha === 'RIGHT') styles.push('text-align:right');

      // 折り返し
      const wrap = fmt.wrapStrategy;
      if (wrap === 'WRAP') styles.push('white-space:normal');

      return `<td style="${styles.join(';')}">${escapeHtml(value)}</td>`;
    }).join('');

    return `<tr>${tds}</tr>`;
  }).join('\n');

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body  { margin:0; padding:6px; background:#fff; font-family:"Noto Sans JP",Arial,sans-serif; }
  table { border-collapse:collapse; font-size:10px; }
  td    { border:1px solid #bbb; padding:2px 4px; min-width:28px; white-space:nowrap; vertical-align:middle; }
</style>
</head>
<body>
<table>${tableRows}</table>
</body>
</html>`;
}

// ── puppeteer でスクリーンショット ────────────────────────
async function screenshotCells(spreadsheetId) {
  const gridData = await fetchCellsData(spreadsheetId);
  const html     = buildHtml(gridData);

  if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });

  const filename   = `sheet_${Date.now()}.png`;
  const outputPath = path.join(TEMP_DIR, filename);

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const tableEl = await page.$('table');
    if (!tableEl) throw new Error('テーブル要素が見つかりません');
    await tableEl.screenshot({ path: outputPath });
  } finally {
    await browser.close();
  }

  // 1時間後に一時ファイルを削除
  setTimeout(() => fs.unlink(outputPath, () => {}), 60 * 60 * 1000);

  return filename;
}

// ── Push 送信先を取得（グループ or ユーザー） ─────────────
function getPushTarget(event) {
  if (event.source?.groupId)  return event.source.groupId;
  if (event.source?.roomId)   return event.source.roomId;
  if (event.source?.userId)   return event.source.userId;
  return process.env.LINE_GROUP_ID || process.env.LINE_USER_ID;
}

// ── バックグラウンドでスクショしてPush送信 ────────────────
async function processAndPush(target, client) {
  try {
    const date          = getSheetBusinessDate();
    const spreadsheetId = await findSpreadsheetId(date);
    const filename      = await screenshotCells(spreadsheetId);

    const baseUrl  = (process.env.LINE_BOT_SERVER_URL || '').replace(/\/$/, '');
    const imageUrl = `${baseUrl}/temp/${filename}`;

    console.log(`[LINE] Push送信: ${imageUrl} → ${target}`);

    await client.pushMessage({
      to: target,
      messages: [{
        type: 'image',
        originalContentUrl: imageUrl,
        previewImageUrl:    imageUrl,
      }],
    });

  } catch (err) {
    console.error('[LINE] Push エラー:', err.message);
    await client.pushMessage({
      to: target,
      messages: [{ type: 'text', text: `エラー: ${err.message}` }],
    }).catch(() => {});
  }
}

// ── LINE イベントハンドラ（200を即返してバックグラウンド処理） ──
function handleLineEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') return;

  const text = event.message.text.trim();
  if (!text.includes('(教えて)') && !text.includes('（教えて）')) return;

  const client = createLineClient();
  const target = getPushTarget(event);

  console.log(`[LINE] (教えて) 受信 → バックグラウンド処理開始 target=${target}`);

  // 即座にreturn（200をLINEに返す）し、裏でスクショ＆Push
  setImmediate(() => processAndPush(target, client));
}

module.exports = { lineConfig, handleLineEvent, middleware };
