/**
 * line_handler.js
 * LINEで「(教えて)」と送ると、当日の日報スプレッドシートCA3:CR32のスクショを返す
 *
 * 日付ルール: 朝6時〜翌5時59分 = 同日（29時制）
 */
require('dotenv').config();
const { middleware, messagingApi } = require('@line/bot-sdk');
const { google } = require('googleapis');
const { createCanvas, registerFont } = require('canvas');
const fs = require('fs');
const path = require('path');

// リポジトリ同梱の日本語フォントを登録
const FONT_PATH = path.join(__dirname, 'fonts', 'NotoSansJP.ttf');
registerFont(FONT_PATH, { family: 'NotoJP' });

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

// ── canvas で PNG 生成 ────────────────────────────────────
function toRgba(c, fallback = '#ffffff') {
  if (!c) return fallback;
  const r = Math.round((c.red   || 0) * 255);
  const g = Math.round((c.green || 0) * 255);
  const b = Math.round((c.blue  || 0) * 255);
  return `rgb(${r},${g},${b})`;
}

async function screenshotCells(spreadsheetId) {
  const gridData = await fetchCellsData(spreadsheetId);
  const rows     = gridData?.rowData || [];
  const colMeta  = gridData?.columnMetadata || [];
  const rowMeta  = gridData?.rowMetadata || [];
  const numCols  = rows.reduce((m, r) => Math.max(m, (r.values || []).length), 0);

  const colWidths  = Array.from({ length: numCols }, (_, i) =>
    colMeta[i]?.pixelSize ? Math.max(colMeta[i].pixelSize * 0.85, 28) : 56);
  const rowHeights = rows.map((_, i) =>
    rowMeta[i]?.pixelSize ? Math.max(rowMeta[i].pixelSize * 0.85, 17) : 20);

  const totalW = colWidths.reduce((s, w) => s + w, 0);
  const totalH = rowHeights.reduce((s, h) => s + h, 0);

  const canvas = createCanvas(totalW, totalH);
  const ctx    = canvas.getContext('2d');

  // 背景
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, totalW, totalH);

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

      // 背景色
      const bg = fmt.backgroundColor;
      const isWhite = !bg || (bg.red===1 && bg.green===1 && bg.blue===1) || (!bg.red && !bg.green && !bg.blue);
      if (!isWhite) {
        ctx.fillStyle = toRgba(bg);
        ctx.fillRect(x, y, w, h);
      }

      // 枠線
      ctx.strokeStyle = '#cccccc';
      ctx.lineWidth = 0.5;
      ctx.strokeRect(x + 0.25, y + 0.25, w - 0.5, h - 0.5);

      // テキスト
      if (value) {
        const fontSize = tf.fontSize ? Math.round(tf.fontSize * 0.82) : 9;
        const bold     = tf.bold ? 'bold ' : '';
        ctx.font = `${bold}${fontSize}px "NotoJP"`;
        ctx.fillStyle = tf.foregroundColor ? toRgba(tf.foregroundColor, '#000000') : '#000000';

        const ha = fmt.horizontalAlignment;
        let tx;
        if (ha === 'CENTER')     { ctx.textAlign = 'center'; tx = x + w / 2; }
        else if (ha === 'RIGHT') { ctx.textAlign = 'right';  tx = x + w - 2; }
        else                     { ctx.textAlign = 'left';   tx = x + 2; }

        const ty = y + h / 2 + fontSize * 0.36;

        // クリップして描画
        ctx.save();
        ctx.beginPath();
        ctx.rect(x, y, w, h);
        ctx.clip();
        ctx.fillText(value, tx, ty);
        ctx.restore();
      }
      x += w;
    }
    y += rowHeights[ri];
  }

  if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
  const filename   = `sheet_${Date.now()}.png`;
  const outputPath = path.join(TEMP_DIR, filename);

  await new Promise((resolve, reject) => {
    const out    = fs.createWriteStream(outputPath);
    const stream = canvas.createPNGStream();
    stream.pipe(out);
    out.on('finish', resolve);
    out.on('error', reject);
  });

  setTimeout(() => fs.unlink(outputPath, () => {}), 60 * 60 * 1000);
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
