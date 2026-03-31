/**
 * line_handler.js
 * LINEで「(教えて)」と送ると、当日の日報スプレッドシートCA3:CR32のスクショを返す
 *
 * 日付ルール: 朝6時〜翌5時59分 = 同日（29時制）
 */
require('dotenv').config();
const { middleware, messagingApi } = require('@line/bot-sdk');
const { google } = require('googleapis');
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

// リポジトリ同梱の日本語フォント（base64）
const FONT_PATH = path.join(__dirname, 'fonts', 'NotoSansJP.otf');
let _fontBase64 = null;
function getFontBase64() {
  if (!_fontBase64) {
    _fontBase64 = fs.readFileSync(FONT_PATH).toString('base64');
  }
  return _fontBase64;
}

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

// ── Sheets API でセルデータ取得 ───────────────────────────
async function fetchCellsData(spreadsheetId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = meta.data.sheets.find(s => s.properties.sheetId === TARGET_GID);
  if (!sheet) throw new Error(`シートが見つかりません (gid=${TARGET_GID})`);
  const sheetName = sheet.properties.title;
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [`'${sheetName}'!${RANGE}`],
    includeGridData: true,
  });
  return res.data.sheets?.[0]?.data?.[0];
}

// ── セルデータ → SVG（日本語フォント埋め込み） ───────────
function escapeXml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function toRgb(c) {
  if (!c) return '#ffffff';
  return `rgb(${Math.round((c.red||0)*255)},${Math.round((c.green||0)*255)},${Math.round((c.blue||0)*255)})`;
}

function buildSvg(gridData) {
  const rows    = gridData?.rowData || [];
  const colMeta = gridData?.columnMetadata || [];
  const rowMeta = gridData?.rowMetadata || [];
  const numCols = rows.reduce((m,r) => Math.max(m,(r.values||[]).length), 0);

  const colWidths  = Array.from({length:numCols}, (_,i) => colMeta[i]?.pixelSize ? Math.max(colMeta[i].pixelSize*0.8,28) : 56);
  const rowHeights = rows.map((_,i) => rowMeta[i]?.pixelSize ? Math.max(rowMeta[i].pixelSize*0.8,17) : 20);
  const totalW = colWidths.reduce((s,w)=>s+w,0);
  const totalH = rowHeights.reduce((s,h)=>s+h,0);

  const fontB64 = getFontBase64();
  let defs = `<defs><style>@font-face{font-family:'NotoJP';src:url('data:font/otf;base64,${fontB64}') format('opentype');}</style>`;
  let rects = '', texts = '';
  let y = 0;

  for (let ri = 0; ri < rows.length; ri++) {
    const cells = rows[ri].values || [];
    let x = 0;
    for (let ci = 0; ci < numCols; ci++) {
      const cell  = cells[ci] || {};
      const value = cell.formattedValue ?? '';
      const fmt   = cell.effectiveFormat || {};
      const tf    = fmt.textFormat || {};
      const w = colWidths[ci], h = rowHeights[ri];

      const bg = fmt.backgroundColor;
      const bgColor = (bg && !(bg.red===1&&bg.green===1&&bg.blue===1)) ? toRgb(bg) : '#ffffff';
      rects += `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="${bgColor}" stroke="#cccccc" stroke-width="0.4"/>`;

      if (value) {
        const color    = tf.foregroundColor ? toRgb(tf.foregroundColor) : '#000000';
        const bold     = tf.bold ? 'bold' : 'normal';
        const fontSize = tf.fontSize ? tf.fontSize * 0.78 : 9;
        const ha       = fmt.horizontalAlignment;
        let tx, anchor;
        if (ha==='CENTER')      { tx=x+w/2; anchor='middle'; }
        else if (ha==='RIGHT')  { tx=x+w-2; anchor='end'; }
        else                    { tx=x+2;   anchor='start'; }
        const ty = y + h/2 + fontSize*0.35;
        const clipId = `c${ri}_${ci}`;
        defs  += `<clipPath id="${clipId}"><rect x="${x}" y="${y}" width="${w}" height="${h}"/></clipPath>`;
        texts += `<text x="${tx}" y="${ty}" font-size="${fontSize}" font-weight="${bold}" fill="${color}" text-anchor="${anchor}" font-family="NotoJP,Arial,sans-serif" clip-path="url(#${clipId})">${escapeXml(value)}</text>`;
      }
      x += w;
    }
    y += rowHeights[ri];
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${totalW}" height="${totalH}">${defs}</defs><rect width="${totalW}" height="${totalH}" fill="#fff"/>${rects}${texts}</svg>`;
}

// ── sharp で SVG → PNG ───────────────────────────────────
async function screenshotCells(spreadsheetId) {
  const gridData = await fetchCellsData(spreadsheetId);
  const svg      = buildSvg(gridData);
  if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
  const filename   = `sheet_${Date.now()}.png`;
  const outputPath = path.join(TEMP_DIR, filename);
  await sharp(Buffer.from(svg)).png().toFile(outputPath);
  setTimeout(() => fs.unlink(outputPath, () => {}), 60*60*1000);
  return filename;
}

// ── Push 送信先: スクショは必ず個人（userId）に送る ──────
function getPushTarget(event) {
  // イベント送信者のuserIdを優先、なければ環境変数のUSER_IDを使用
  if (event.source?.userId) return event.source.userId;
  return process.env.LINE_USER_ID;
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
  if (!text.includes('教えて')) return;

  const client = createLineClient();
  const target = getPushTarget(event);

  console.log(`[LINE] (教えて) 受信 → バックグラウンド処理開始 target=${target}`);

  // 即座にreturn（200をLINEに返す）し、裏でスクショ＆Push
  setImmediate(() => processAndPush(target, client));
}

module.exports = { lineConfig, handleLineEvent, middleware };
