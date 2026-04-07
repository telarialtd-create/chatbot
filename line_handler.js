/**
 * line_handler.js
 * ・「教えて」→ 当日の日報 CA3:CR32 スクショを個人に送信
 * ・「【日付、名前】」→ 指定日の日報の明細シートH5に名前を入力し
 *   setMeisaiFromUriage スクリプトを実行、H2:M25 スクショをグループに送信
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

// ── 明細機能の定数 ────────────────────────────────────────
const MEISAI_SHEET_NAME = '明細';
const MEISAI_RANGE = 'H2:M25';
const MEISAI_NAME_CELL = 'H5';
const MEISAI_SCRIPT_FUNCTION = 'setMeisaiFromUriage';

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
// サーバーはUTC → JST(+9h)に変換してから判定
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

// ── Drive からスプレッドシート ID を取得 ──────────────────
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
  console.log(`[LINE] 検索結果 (${files.length}件):`, files.map(f => f.name).join(', '));
  const file = files[0];
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

// ── 明細機能 ──────────────────────────────────────────────

// 以下の形式をすべてサポート:
//   【4月7日、ここあ】  【4月7日,ここあ】
//   4月7日 ここあ  4月7日　ここあ（全角スペース）
//   複数行形式:
//     日時　4月7日
//     名前　ここあ
//     交通費　片道      ← 省略可（省略時はH18の値をそのまま使用）
function parseMeisaiRequest(text) {
  // 複数行フォーマット（日時・名前・交通費）
  if (/日時/.test(text) && /名前/.test(text)) {
    const dateMatch      = text.match(/日時[\s　:：]*(.+)/);
    const nameMatch      = text.match(/名前[\s　:：]*(.+)/);
    const transportMatch = text.match(/交通費[\s　:：]*(.+)/);
    if (dateMatch && nameMatch) {
      const transportRaw = transportMatch ? transportMatch[1].trim() : null;
      // 片道/往復/P のいずれかに正規化（それ以外は null）
      const transportType = normalizeTransportType(transportRaw);
      return {
        dateStr:       dateMatch[1].trim(),
        name:          nameMatch[1].trim(),
        transportType,
      };
    }
  }

  // 【】形式
  const bracketMatch = text.match(/【(.+?)[,、](.+?)】/);
  if (bracketMatch) return { dateStr: bracketMatch[1].trim(), name: bracketMatch[2].trim(), transportType: null };

  // 「日付 名前」スペース区切り形式（日付は〇月〇日 を含む）
  const spaceMatch = text.match(/^((?:\d{4}年)?\d{1,2}月\d{1,2}日)[\s　]+(\S+)$/);
  if (spaceMatch) return { dateStr: spaceMatch[1].trim(), name: spaceMatch[2].trim(), transportType: null };

  return null;
}

// 交通費種別を正規化（片道/往復/P → そのまま返す、それ以外はnull）
function normalizeTransportType(raw) {
  if (!raw) return null;
  const s = raw.trim();
  if (s === '片道') return '片道';
  if (s === '往復') return '往復';
  if (s.toUpperCase() === 'P' || s === 'ｐ') return 'P';
  return null;
}

// 日付文字列でファイルを検索（部分一致）
// MEISAI_TEST_ID が設定されている場合はそのスプレッドシートを使用
async function findSpreadsheetByDateStr(dateStr) {
  if (process.env.MEISAI_TEST_ID) {
    console.log(`[明細] テストモード: MEISAI_TEST_ID=${process.env.MEISAI_TEST_ID}`);
    return process.env.MEISAI_TEST_ID;
  }
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const res = await drive.files.list({
    q: `'${NIPPO_FOLDER_ID}' in parents and name contains '${dateStr}'`,
    fields: 'files(id, name)',
    orderBy: 'createdTime desc',
    pageSize: 5,
  });
  const files = res.data.files || [];
  console.log(`[明細] 検索結果 (${files.length}件):`, files.map(f => f.name).join(', '));
  const file = files[0];
  if (!file) throw new Error(`ファイルが見つかりません: ${dateStr}`);
  console.log(`[明細] ファイル発見: ${file.name} (${file.id})`);
  return file.id;
}

// 指定セルに値を書き込む
async function writeCellValue(spreadsheetId, sheetName, cell, value) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${sheetName}'!${cell}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[value]] },
  });
  console.log(`[明細] ${sheetName}!${cell} に書き込み完了: ${value}`);
}

// setMeisaiFromUriage 相当の処理を Node.js で直接実装
// 売上シートから名前でフィルタし、明細シートに書き込む
async function runAppsScript(spreadsheetId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // --- 1. 明細!H5:K5 から源氏名を取得 ---
  const nameRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${MEISAI_SHEET_NAME}'!H5:K5`,
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const nameRow = (nameRes.data.values || [[]])[0] || [];
  const genjimei = nameRow.find(v => v && String(v).trim() !== '') || '';
  if (!genjimei) {
    console.log('[明細] 源氏名が未設定のためスキップ');
    return;
  }
  console.log(`[明細] 源氏名: ${genjimei}`);

  // --- 2. 明細!H18 から交通費種別を取得 ---
  const transportRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${MEISAI_SHEET_NAME}'!H18`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const transportType = ((transportRes.data.values || [[]])[0] || [])[0] || '';

  // --- 3. マスターシートから交通費を取得 ---
  const MASTER_SS_ID = '1sh6MgPL4k2StMofKuys_5QA0-3JiDbdY98qe8bvAJEo';
  let baseTransportFee = 0;
  try {
    const masterRes = await sheets.spreadsheets.values.get({
      spreadsheetId: MASTER_SS_ID,
      range: 'E:F',
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const masterRows = masterRes.data.values || [];
    for (const row of masterRows) {
      if (row[0] === genjimei) { baseTransportFee = Number(row[1]) || 0; break; }
    }
  } catch (e) {
    console.log('[明細] マスターシート取得エラー（交通費=0）:', e.message);
  }
  const finalTransportFee =
    transportType === '片道' ? baseTransportFee :
    transportType === '往復' ? baseTransportFee * 2 : 0;

  // --- 4. 前回分を取得（明細!Q30:R500 または T30:U500） ---
  let zenkaiValue = 0;
  try {
    const z1Res = await sheets.spreadsheets.values.get({
      spreadsheetId, range: `'${MEISAI_SHEET_NAME}'!Q30:R500`, valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const z2Res = await sheets.spreadsheets.values.get({
      spreadsheetId, range: `'${MEISAI_SHEET_NAME}'!T30:U500`, valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const findZenkai = (rows) => {
      for (const r of (rows || [])) { if (r[0] === genjimei) return Number(r[1]) || 0; }
      return null;
    };
    zenkaiValue = findZenkai(z1Res.data.values) ?? findZenkai(z2Res.data.values) ?? 0;
  } catch (e) {
    console.log('[明細] 前回分取得エラー（=0）:', e.message);
  }

  // --- 5. 売上シートから該当名のデータを取得（AN:BB列） ---
  // AN=本数, AO=店名, AP=源氏名, AQ=指名, AR=OP, AV=コース, AZ=料金(×1000), BB=給料(×1000)
  const urRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'売上'!AN:BB",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const urRows = (urRes.data.values || []).slice(1); // ヘッダー除く
  const toK = v => (v === '' || v === null || v === undefined || isNaN(Number(v))) ? '' : Number(v) * 1000;
  const meisaiRows = [];
  let cashlessTotal = 0;

  for (const row of urRows) {
    if (row[2] !== genjimei) continue; // AP列=源氏名
    const honsu   = row[0]  ?? '';  // AN=本数（0も有効値なので ?? を使う）
    const course  = row[8]  ?? '';  // AV=コース
    const shimei  = row[3]  ?? '';  // AQ=指名
    const ryokin  = toK(row[12]);   // AZ=料金
    const kyuryo  = toK(row[14]);   // BB=給料
    const option  = row[4]  ?? '';  // AR=OP
    meisaiRows.push([honsu, course, shimei, ryokin, kyuryo, option]);
    if (/カード|ペイペイ|paypay/i.test(String(option)) && ryokin !== '') {
      cashlessTotal += Number(ryokin);
    }
  }
  console.log(`[明細] 売上データ: ${meisaiRows.length}件`);

  // --- 6. 明細シートへの書き込み ---
  const batchData = [];

  // H8:M15 をクリアしてから書き込み
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H8:M15`, values: Array(8).fill(Array(6).fill('')) });
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H36:M55`, values: Array(20).fill(Array(6).fill('')) });
  if (meisaiRows.length > 0) {
    batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H8`, values: meisaiRows });
    batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H36`, values: meisaiRows });
  }
  // 交通費・前回分・キャッシュレス合計（L21は毎回上書き）
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H21`, values: [[finalTransportFee]] });
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!K21`, values: [[zenkaiValue]] });
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!L21`, values: [[cashlessTotal > 0 ? cashlessTotal : '']] });

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data: batchData },
  });
  console.log('[明細] setMeisaiFromUriage 完了（交通費:', finalTransportFee, '前回分:', zenkaiValue, 'キャッシュレス:', cashlessTotal, '）');
}

// 明細シート H2:M25 をスクショ（結合セル対応）
async function screenshotMeisai(spreadsheetId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [`'${MEISAI_SHEET_NAME}'!${MEISAI_RANGE}`],
    includeGridData: true,
  });

  const sheetData = res.data.sheets?.[0];
  const gridData  = sheetData?.data?.[0];
  const allMerges = sheetData?.merges || [];

  const rows    = gridData?.rowData || [];
  const colMeta = gridData?.columnMetadata || [];
  const rowMeta = gridData?.rowMetadata || [];
  const numCols = rows.reduce((m, r) => Math.max(m, (r.values || []).length), 0);

  // H2:M25 のシートグローバル座標オフセット（0-indexed）
  // H = 列8(1-based) = 列7(0-based) / 行2(1-based) = 行1(0-based)
  const ROW_OFFSET = 1;
  const COL_OFFSET = 7;

  // 結合セルマップを構築
  // mergeMap["ri,ci"] = { rowSpan, colSpan }  ← 結合の左上セル
  // mergedSet          = 結合に含まれる「左上以外」のセル集合
  const mergeMap  = {};
  const mergedSet = new Set();

  for (const m of allMerges) {
    const r0 = m.startRowIndex    - ROW_OFFSET;
    const c0 = m.startColumnIndex - COL_OFFSET;
    const r1 = m.endRowIndex      - ROW_OFFSET;
    const c1 = m.endColumnIndex   - COL_OFFSET;
    if (r1 <= 0 || c1 <= 0 || r0 >= rows.length || c0 >= numCols) continue;
    const sr = Math.max(r0, 0);
    const sc = Math.max(c0, 0);
    mergeMap[`${sr},${sc}`] = { rowSpan: r1 - sr, colSpan: c1 - sc };
    for (let r = sr; r < r1; r++) {
      for (let c = sc; c < c1; c++) {
        if (r === sr && c === sc) continue;
        mergedSet.add(`${r},${c}`);
      }
    }
  }

  const colWidths  = Array.from({ length: numCols }, (_, i) =>
    colMeta[i]?.pixelSize ? Math.max(colMeta[i].pixelSize * 0.85, 28) : 56);
  const rowHeights = rows.map((_, i) =>
    rowMeta[i]?.pixelSize ? Math.max(rowMeta[i].pixelSize * 0.85, 17) : 20);

  const totalW = colWidths.reduce((s, w) => s + w, 0);
  const totalH = rowHeights.reduce((s, h) => s + h, 0);
  const canvas = createCanvas(totalW, totalH);
  const ctx    = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, totalW, totalH);

  let y = 0;
  for (let ri = 0; ri < rows.length; ri++) {
    const cells = rows[ri].values || [];
    let x = 0;
    for (let ci = 0; ci < numCols; ci++) {
      const w = colWidths[ci];
      const h = rowHeights[ri];

      if (mergedSet.has(`${ri},${ci}`)) {
        // 結合の左上以外のセル → 描画スキップ（背景・枠線とも左上セル側で描画済み）
        x += w;
        continue;
      }

      const cell  = cells[ci] || {};
      const fmt   = cell.effectiveFormat || {};
      const tf    = fmt.textFormat || {};

      // 結合セルの左上なら結合範囲分の幅・高さを計算
      const mergeInfo = mergeMap[`${ri},${ci}`];
      let cellW = w, cellH = h;
      if (mergeInfo) {
        for (let dc = 1; dc < mergeInfo.colSpan; dc++) cellW += (colWidths[ci + dc] || 0);
        for (let dr = 1; dr < mergeInfo.rowSpan; dr++) cellH += (rowHeights[ri + dr] || 0);
      }

      // 背景色
      const bg = fmt.backgroundColor;
      const isWhite = !bg || (bg.red === 1 && bg.green === 1 && bg.blue === 1) || (!bg.red && !bg.green && !bg.blue);
      if (!isWhite) {
        ctx.fillStyle = toRgba(bg);
        ctx.fillRect(x, y, cellW, cellH);
      }

      // 枠線
      ctx.strokeStyle = '#cccccc';
      ctx.lineWidth   = 0.5;
      ctx.strokeRect(x + 0.25, y + 0.25, cellW - 0.5, cellH - 0.5);

      // テキスト（チェックボックス等のbool値はスキップ）
      const isBoolean = cell.userEnteredValue?.boolValue !== undefined;
      const value     = isBoolean ? '' : (cell.formattedValue ?? '');
      if (value) {
        const fontSize = tf.fontSize ? Math.round(tf.fontSize * 0.82) : 9;
        const bold     = tf.bold ? 'bold ' : '';
        ctx.font       = `${bold}${fontSize}px "NotoJP"`;
        ctx.fillStyle  = tf.foregroundColor ? toRgba(tf.foregroundColor, '#000000') : '#000000';
        const ha = fmt.horizontalAlignment;
        let tx;
        if (ha === 'CENTER')     { ctx.textAlign = 'center'; tx = x + cellW / 2; }
        else if (ha === 'RIGHT') { ctx.textAlign = 'right';  tx = x + cellW - 2; }
        else                     { ctx.textAlign = 'left';   tx = x + 2; }
        const ty = y + cellH / 2 + fontSize * 0.36;
        ctx.save();
        ctx.beginPath();
        ctx.rect(x, y, cellW, cellH);
        ctx.clip();
        ctx.fillText(value, tx, ty);
        ctx.restore();
      }
      x += w;
    }
    y += rowHeights[ri];
  }

  if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
  const filename   = `meisai_${Date.now()}.png`;
  const outputPath = path.join(TEMP_DIR, filename);
  await new Promise((resolve, reject) => {
    const out = fs.createWriteStream(outputPath);
    canvas.createPNGStream().pipe(out);
    out.on('finish', resolve);
    out.on('error', reject);
  });
  setTimeout(() => fs.unlink(outputPath, () => {}), 60 * 60 * 1000);
  return filename;
}

// 明細処理本体（バックグラウンド）
async function processMeisaiAndPush(target, client, dateStr, name, transportType = null) {
  try {
    const spreadsheetId = await findSpreadsheetByDateStr(dateStr);
    await writeCellValue(spreadsheetId, MEISAI_SHEET_NAME, MEISAI_NAME_CELL, name);
    // 交通費種別が指定されていればH18に書き込む（省略時はシートの既存値を使用）
    if (transportType !== null) {
      await writeCellValue(spreadsheetId, MEISAI_SHEET_NAME, 'H18', transportType);
    }
    await runAppsScript(spreadsheetId);
    // スクリプト完了を3秒待つ
    await new Promise(r => setTimeout(r, 3000));
    const filename = await screenshotMeisai(spreadsheetId);
    const baseUrl = (process.env.LINE_BOT_SERVER_URL || '').replace(/\/$/, '');
    const imageUrl = `${baseUrl}/temp/${filename}`;
    console.log(`[明細] Push送信: ${imageUrl} → ${target}`);
    await client.pushMessage({
      to: target,
      messages: [{ type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl }],
    });
  } catch (err) {
    console.error('[明細] エラー:', err.message);
    await client.pushMessage({
      to: target,
      messages: [{ type: 'text', text: `明細エラー: ${err.message}` }],
    }).catch(() => {});
  }
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
    console.log(`[LINE] 検索日付(JST): ${dateToNippoName(date)}`);
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
  const client = createLineClient();

  // 【日付、名前】形式: 明細スクショ
  const meisai = parseMeisaiRequest(text);
  if (meisai) {
    const target = event.source?.groupId || event.source?.userId || process.env.LINE_USER_ID;
    console.log(`[LINE] 明細リクエスト受信: ${meisai.dateStr} / ${meisai.name} / 交通費=${meisai.transportType ?? '指定なし'} → target=${target}`);
    setImmediate(() => processMeisaiAndPush(target, client, meisai.dateStr, meisai.name, meisai.transportType));
    return;
  }

  // 「教えて」: 日報スクショ（既存機能）
  if (text.includes('教えて')) {
    const target = getPushTarget(event);
    console.log(`[LINE] (教えて) 受信 → バックグラウンド処理開始 target=${target}`);
    setImmediate(() => processAndPush(target, client));
    return;
  }
}

module.exports = { lineConfig, handleLineEvent, middleware };
