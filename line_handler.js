/**
 * line_handler.js
 * ・「教えて」→ 当日の日報 CA3:CR32 スクショを個人に送信
 * ・「【日付、名前】」→ 指定日の日報の明細シートH5に名前を入力し
 *   setMeisaiFromUriage スクリプトを実行、H2:M25 スクショをグループに送信
 * ・「月報更新」→ 昨日の日報データを月報に反映（手動トリガー）
 * ・「月報更新 4月7日」→ 指定日の日報を月報に反映
 * ・電話番号（0から始まる10〜11桁）→ 利用履歴シートを検索して最新30件を返信
 */
require('dotenv').config();
const { middleware, messagingApi } = require('@line/bot-sdk');
const { syncNippoToGeppo } = require('./nippo_to_geppo_v2');
const { isNippoInput, processNippoInput, getStoresByUserId, getFolderByStoreId } = require('./line_nippo_input');
const { google } = require('googleapis');
const { createCanvas, registerFont } = require('canvas');
const fs = require('fs');
const path = require('path');

// リポジトリ同梱の日本語フォントを登録
const FONT_PATH = path.join(__dirname, 'fonts', 'NotoSansJP.ttf');
registerFont(FONT_PATH, { family: 'NotoJP' });

const NIPPO_FOLDER_ID = '16R1BK5NnvYkH4Eqh6t51OGQl0tXVJ3If';
const TARGET_GID = 1873674341;
const RANGE = 'CA3:CR32';
const TEMP_DIR = path.join(__dirname, 'public', 'temp');

// ── 明細機能の定数 ────────────────────────────────────────
const MEISAI_SHEET_NAME = '明細';
const MEISAI_RANGE = 'H2:M25';
const MEISAI_NAME_CELL = 'H5';
const MEISAI_SCRIPT_FUNCTION = 'setMeisaiFromUriage';

// ── 明細書（新フォーマット）の定数 ───────────────────────
// C3=名前 / D40=交通費 / E41=バンス / E46=お釣り / E48=転記チェック
// 描画範囲 A1:G47 を PNG で返信
const MEISAISYO_SHEET_NAME = '明細書';
const MEISAISYO_RANGE = 'A1:G47';
const MEISAISYO_NAME_CELL = 'C3';
const MEISAISYO_TRANSPORT_CELL = 'D40';
const MEISAISYO_BANCE_CELL = 'E41';
const MEISAISYO_OTSURI_CELL = 'E46';
const MEISAISYO_CHECK_CELL = 'E48';

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
  let client_id, client_secret, refresh_token;

  const credsPath = path.join(process.env.HOME, '.config/gdrive-server-credentials.json');
  const keysPath  = path.join(process.env.HOME, '.config/gcp-oauth.keys.json');

  if (fs.existsSync(credsPath) && fs.existsSync(keysPath)) {
    // 認証ファイルが存在する場合はファイルを優先（access_tokenは使わない：期限切れになるため）
    const oauthKeys   = JSON.parse(fs.readFileSync(keysPath));
    const credentials = JSON.parse(fs.readFileSync(credsPath));
    const installed   = oauthKeys.installed || oauthKeys.web;
    client_id     = installed.client_id;
    client_secret = installed.client_secret;
    refresh_token = credentials.refresh_token;
  } else {
    // ファイルがない場合のみ環境変数を使用
    client_id     = process.env.GOOGLE_CLIENT_ID;
    client_secret = process.env.GOOGLE_CLIENT_SECRET;
    refresh_token = process.env.GOOGLE_REFRESH_TOKEN;
  }

  const client = new google.auth.OAuth2(client_id, client_secret, 'http://localhost');
  client.setCredentials({ refresh_token }); // access_tokenは設定しない（期限切れ防止）
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

// 半角カタカナ → 全角カタカナ変換（NotoJPフォントで正常描画するため）
// 例: ｺｰｽ → コース
function toFullWidth(str) {
  // 半角カタカナの基本マッピング（濁点・半濁点付きは2文字→1文字に合成）
  const map = {
    'ｦ':'ヲ','ｧ':'ァ','ｨ':'ィ','ｩ':'ゥ','ｪ':'ェ','ｫ':'ォ','ｬ':'ャ','ｭ':'ュ','ｮ':'ョ','ｯ':'ッ',
    'ｰ':'ー','ｱ':'ア','ｲ':'イ','ｳ':'ウ','ｴ':'エ','ｵ':'オ','ｶ':'カ','ｷ':'キ','ｸ':'ク','ｹ':'ケ',
    'ｺ':'コ','ｻ':'サ','ｼ':'シ','ｽ':'ス','ｾ':'セ','ｿ':'ソ','ﾀ':'タ','ﾁ':'チ','ﾂ':'ツ','ﾃ':'テ',
    'ﾄ':'ト','ﾅ':'ナ','ﾆ':'ニ','ﾇ':'ヌ','ﾈ':'ネ','ﾉ':'ノ','ﾊ':'ハ','ﾋ':'ヒ','ﾌ':'フ','ﾍ':'ヘ',
    'ﾎ':'ホ','ﾏ':'マ','ﾐ':'ミ','ﾑ':'ム','ﾒ':'メ','ﾓ':'モ','ﾔ':'ヤ','ﾕ':'ユ','ﾖ':'ヨ','ﾗ':'ラ',
    'ﾘ':'リ','ﾙ':'ル','ﾚ':'レ','ﾛ':'ロ','ﾜ':'ワ','ｦ':'ヲ','ﾝ':'ン',
  };
  // 濁点・半濁点付き合成 (例: ｶﾞ→ガ)
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
    if (next === 'ﾞ' && dakuMap[ch])    { result += dakuMap[ch];    i++; }
    else if (next === 'ﾟ' && handakuMap[ch]) { result += handakuMap[ch]; i++; }
    else                                 { result += ch; }
  }
  return result;
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

  const totalW = Math.round(colWidths.reduce((s, w) => s + w, 0));
  const totalH = Math.round(rowHeights.reduce((s, h) => s + h, 0));

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
      const value = toFullWidth(cell.formattedValue ?? '');
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

// 「3,000」「3000円」「￥3000」などを整数に正規化（パース不可は null）
function parseAmount(raw) {
  if (raw === null || raw === undefined) return null;
  const cleaned = String(raw).replace(/[,，\s　円¥￥]/g, '');
  if (!cleaned) return null;
  const n = parseInt(cleaned, 10);
  return Number.isNaN(n) ? null : n;
}

// 「4/27」「2026/4/27」「4-27」を「4月27日」へ正規化（既に「月日」形式なら素通し）
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

// 明細書（新フォーマット）リクエストの解析
// `【明細】` と `T-XXX` の両方を含む複数行メッセージから店舗ID/日付/名前/交通費/バンス/お釣りを抽出する。
// 例:
//   T-001
//   【明細】
//   日付:4/27
//   名前:柚月
//   交通費:片道
//   バンス:3000
//   お釣り:1000
const MEISAISYO_TRIGGER = /[【\[]\s*明\s*細\s*[】\]]/;
function parseMeisaisyoRequest(text) {
  if (!text) return null;
  if (!MEISAISYO_TRIGGER.test(text)) return null;

  // 全角→半角 normalize（T-XXX判定のため）
  const normalized = String(text)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[－ー―−–—]/g, '-');

  // T-XXX を独立行として抽出（例: "T-001", "T1043" などを許容）
  let storeId = null;
  for (const rawLine of normalized.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (/^[Tt]-?\d{3,5}$/.test(line)) {
      storeId = line.toUpperCase();
      if (!storeId.includes('-')) storeId = storeId.replace(/^T(\d+)/, 'T-$1');
      break;
    }
  }
  if (!storeId) return null;

  let dateStr = null, name = null, transport = null, bance = null, otsuri = null;
  // トリガー文字列を除去（行頭/行末に単独で書かれていても、文中に書かれていても削る）
  const cleaned = text.replace(MEISAISYO_TRIGGER, '');
  const lines = cleaned.split(/\r?\n/).map(l => l.trim()).filter(Boolean);

  for (const line of lines) {
    let m;
    if ((m = line.match(/^(?:日付|日時)[\s　:：]*(.*)$/))) { if (!dateStr && m[1]) dateStr = normalizeSlashDate(m[1].trim()); continue; }
    if ((m = line.match(/^名前[\s　:：]*(.*)$/)))         { if (!name && m[1]) name = m[1].trim(); continue; }
    if ((m = line.match(/^交通費[\s　:：]*(.*)$/)))       { if (m[1]) transport = normalizeTransportType(m[1]); continue; }
    if ((m = line.match(/^バンス[\s　:：]*(.*)$/)))       { if (m[1]) bance = parseAmount(m[1]); continue; }
    if ((m = line.match(/^お釣り[\s　:：]*(.*)$/)))       { if (m[1]) otsuri = parseAmount(m[1]); continue; }
  }

  if (!dateStr || !name) return null;
  return { storeId, dateStr, name, transport, bance, otsuri };
}

// 交通費種別を正規化（片道/往復/P/なし → そのまま返す、それ以外はnull）
function normalizeTransportType(raw) {
  if (!raw) return null;
  const s = raw.trim();
  if (s === '片道') return '片道';
  if (s === '往復') return '往復';
  if (s.toUpperCase() === 'P' || s === 'ｐ') return 'P';
  if (s === 'なし' || s === 'ナシ' || s === 'なし') return 'なし';
  return null;
}

// 「4月7日」→「2026年4月7日」のように年を補完する
function normalizeDateStr(dateStr) {
  // すでに年が含まれている場合はそのまま
  if (/\d{4}年/.test(dateStr)) return dateStr;
  // 現在のJST年を取得
  const jst = getJSTDate();
  const year = jst.getUTCFullYear();
  return `${year}年${dateStr}`;
}

// 日付文字列でファイルを検索
// MEISAI_TEST_ID が設定されている場合はそのスプレッドシートを使用
// folderId が指定されない場合は環境変数 NIPPO_FOLDER_ID を使用
async function findSpreadsheetByDateStr(dateStr, folderId = null) {
  if (process.env.MEISAI_TEST_ID) {
    console.log(`[明細] テストモード: MEISAI_TEST_ID=${process.env.MEISAI_TEST_ID}`);
    return process.env.MEISAI_TEST_ID;
  }
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });

  // 年を補完して検索精度を上げる（「4月7日」→「2026年4月7日」）
  const normalizedDate = normalizeDateStr(dateStr);
  const targetFolderId = folderId || NIPPO_FOLDER_ID;
  console.log(`[明細] 検索キー: "${dateStr}" → "${normalizedDate}" / folderId=${targetFolderId}`);

  const res = await drive.files.list({
    q: `'${targetFolderId}' in parents and name contains '${normalizedDate}'`,
    fields: 'files(id, name)',
    orderBy: 'createdTime desc',
    pageSize: 10,
  });
  const files = res.data.files || [];
  console.log(`[明細] 検索結果 (${files.length}件):`, files.map(f => f.name).join(', '));
  if (!files.length) throw new Error(`ファイルが見つかりません: ${normalizedDate}`);

  // 複数ヒット時は名前が短いもの（店舗名サフィックスがない本体ファイル）を優先
  files.sort((a, b) => a.name.length - b.name.length);
  const file = files[0];
  console.log(`[明細] 選択ファイル: ${file.name} (${file.id})`);
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
// A1記法の左上セルから 0-indexed の (row, col) オフセットを計算
function rangeStartOffsets(range) {
  const m = String(range).match(/^([A-Z]+)(\d+)/);
  if (!m) return { rowOffset: 0, colOffset: 0 };
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { rowOffset: parseInt(m[2], 10) - 1, colOffset: col - 1 };
}

async function screenshotMeisai(spreadsheetId, sheetName = MEISAI_SHEET_NAME, range = MEISAI_RANGE, prefix = 'meisai', opts = {}) {
  const { compactEmpty = false } = opts;
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [`'${sheetName}'!${range}`],
    includeGridData: true,
  });

  const sheetData = res.data.sheets?.[0];
  const gridData  = sheetData?.data?.[0];
  const allMerges = sheetData?.merges || [];

  const rows    = gridData?.rowData || [];
  const colMeta = gridData?.columnMetadata || [];
  const rowMeta = gridData?.rowMetadata || [];
  const numCols = rows.reduce((m, r) => Math.max(m, (r.values || []).length), 0);

  // 範囲左上のシートグローバル座標オフセット（0-indexed）
  const { rowOffset: ROW_OFFSET, colOffset: COL_OFFSET } = rangeStartOffsets(range);

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

  // 空白行スキップ判定: 値が一切無く、結合セルにも含まれない行
  const skipRow = new Set();
  if (compactEmpty) {
    for (let ri = 0; ri < rows.length; ri++) {
      const cells = rows[ri].values || [];
      const hasContent = cells.some(c => {
        if (!c) return false;
        if (c.formattedValue !== undefined && c.formattedValue !== '') return true;
        if (c.userEnteredValue?.boolValue !== undefined) return true;
        return false;
      });
      if (hasContent) continue;
      // 結合の途中行は残す（左上以外のセルが当該行にある場合）
      let inMergeBody = false;
      for (let ci = 0; ci < numCols; ci++) {
        if (mergedSet.has(`${ri},${ci}`)) { inMergeBody = true; break; }
        if (mergeMap[`${ri},${ci}`]) { inMergeBody = true; break; }
      }
      if (!inMergeBody) skipRow.add(ri);
    }
  }

  const colWidths  = Array.from({ length: numCols }, (_, i) =>
    colMeta[i]?.pixelSize ? Math.max(colMeta[i].pixelSize * 0.85, 28) : 56);
  const rowHeights = rows.map((_, i) =>
    rowMeta[i]?.pixelSize ? Math.max(rowMeta[i].pixelSize * 0.85, 17) : 20);

  const totalW = Math.round(colWidths.reduce((s, w) => s + w, 0));
  const totalH = Math.round(rows.reduce((s, _, i) => s + (skipRow.has(i) ? 0 : rowHeights[i]), 0));
  const canvas = createCanvas(totalW, totalH);
  const ctx    = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, totalW, totalH);

  let y = 0;
  for (let ri = 0; ri < rows.length; ri++) {
    if (skipRow.has(ri)) continue;
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
      const value     = isBoolean ? '' : toFullWidth(cell.formattedValue ?? '');
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
  const filename   = `${prefix}_${Date.now()}.png`;
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

// 明細書（新フォーマット）処理本体
// parsed: { storeId, dateStr, name, transport, bance, otsuri }
// userId: LINEイベントの発言者userId（許可チェック用）
async function processMeisaisyoAndPush(target, client, parsed, userId) {
  try {
    // T-XXX から店舗フォルダIDを取得（許可userIdチェック付き）
    const folderInfo = await getFolderByStoreId(parsed.storeId, userId);
    console.log(`[明細書] 店舗特定: ${folderInfo.storeId}/${folderInfo.storeName} folderId=${folderInfo.folderId}`);

    const spreadsheetId = await findSpreadsheetByDateStr(parsed.dateStr, folderInfo.folderId);
    const auth = createAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });

    // Phase 1: 名前(C3) を書込し、シート側 onEdit トリガーで消えるはずの E41/E46 を
    //          API書込では発火しないため明示的にクリアする
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_NAME_CELL}`,   values: [[parsed.name]] },
          { range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_BANCE_CELL}`,  values: [['']] },
          { range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_OTSURI_CELL}`, values: [['']] },
        ],
      },
    });
    console.log(`[明細書] 名前書込&E41/E46クリア完了 ssid=${spreadsheetId} name=${parsed.name}`);

    // Phase 2: 残りのフィールドを書込
    const data = [];
    if (parsed.transport !== null && parsed.transport !== undefined) {
      data.push({ range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_TRANSPORT_CELL}`, values: [[parsed.transport]] });
    }
    if (parsed.bance !== null && parsed.bance !== undefined) {
      data.push({ range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_BANCE_CELL}`, values: [[parsed.bance]] });
    }
    if (parsed.otsuri !== null && parsed.otsuri !== undefined) {
      data.push({ range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_OTSURI_CELL}`, values: [[parsed.otsuri]] });
    }
    data.push({ range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_CHECK_CELL}`, values: [[true]] });

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data },
    });
    console.log(`[明細書] 残フィールド書込完了 transport=${parsed.transport} bance=${parsed.bance} otsuri=${parsed.otsuri}`);

    const filename = await screenshotMeisai(spreadsheetId, MEISAISYO_SHEET_NAME, MEISAISYO_RANGE, 'meisaisyo', { compactEmpty: true });
    const baseUrl = (process.env.LINE_BOT_SERVER_URL || '').replace(/\/$/, '');
    const imageUrl = `${baseUrl}/temp/${filename}`;
    console.log(`[明細書] Push送信: ${imageUrl} → ${target}`);
    await client.pushMessage({
      to: target,
      messages: [{ type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl }],
    });
  } catch (err) {
    console.error('[明細書] エラー:', err.message);
    await client.pushMessage({
      to: target,
      messages: [{ type: 'text', text: `明細書エラー: ${err.message}` }],
    }).catch(() => {});
  }
}

// 明細処理本体（バックグラウンド）
async function processMeisaiAndPush(target, client, dateStr, name, transportType = null) {
  try {
    const spreadsheetId = await findSpreadsheetByDateStr(dateStr);
    await writeCellValue(spreadsheetId, MEISAI_SHEET_NAME, MEISAI_NAME_CELL, name);
    // 交通費種別が指定されていればH18に書き込む（省略時はシートの既存値を使用）
    if (transportType !== null) {
      console.log(`[明細] 交通費種別をH18に書き込み: "${transportType}"`);
      await writeCellValue(spreadsheetId, MEISAI_SHEET_NAME, 'H18', transportType);
      console.log(`[明細] H18書き込み完了`);
    } else {
      console.log(`[明細] 交通費省略 → H18はそのまま`);
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

// ── 顧客更新（日報全件 → 利用履歴コピー）────────────────────
const KOKYAKU_SOURCE_SHEET_ID   = '1rPHZ75QnjhDxwqsITFBKgVDvuPVo_dz4lT7L22POw3k';
const KOKYAKU_TARGET_SHEET_ID   = '1A_LaiWm2QhvXk4jKINSRvKKX3-xKYLzygNnlY3ZtI9A';
const KOKYAKU_TARGET_SHEET_NAME = '利用履歴';
const NIPPO_ZENKEN_SHEET_NAME  = '日報_全件';

// 「4月23日 顧客更新」「4月23日　顧客更新」形式をパース
function parseKokyakuUpdateCommand(text) {
  const m = text.match(/^(\d{1,2}月\d{1,2}日)[\s　]+顧客更新$/);
  if (m) return { dateStr: m[1] };
  // 【4月23日　顧客更新】形式も対応
  const m2 = text.match(/^【(\d{1,2}月\d{1,2}日)[\s　]+顧客更新】$/);
  if (m2) return { dateStr: m2[1] };
  return null;
}

// 「2026年4月23日」→「2026/4/23」に変換（利用履歴の既存フォーマットに合わせる）
function toSlashDate(nenGappiStr) {
  const m = nenGappiStr.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (!m) return nenGappiStr;
  return `${m[1]}/${parseInt(m[2])}/${parseInt(m[3])}`;
}

// 日報全件シートから利用履歴シートへコピー
const KOKYAKU_VERSION = 'v3-2026-04-24';
async function processKokyakuUpdate(dateStr) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // 1. ソーススプレッドシート（日付から動的に検索）
  const normalizedDate = normalizeDateStr(dateStr);
  const slashDate = toSlashDate(normalizedDate);
  let spreadsheetId;
  try {
    spreadsheetId = await findSpreadsheetByDateStr(dateStr);
  } catch (e) {
    throw new Error(`[${KOKYAKU_VERSION}][日報検索失敗] folder=${NIPPO_FOLDER_ID} date=${normalizedDate} orig=${e.message}`);
  }
  console.log(`[顧客更新] 日報ファイル: ${spreadsheetId}`);

  // 2. 日報全件シートからB3:J列を読み取り
  let srcRes;
  try {
    srcRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${NIPPO_ZENKEN_SHEET_NAME}'!B3:J`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
  } catch (e) {
    // どのシート/IDで失敗したか特定できるようにする
    const meta = await sheets.spreadsheets.get({ spreadsheetId }).catch(() => null);
    const sheetNames = meta?.data?.sheets?.map(s => s.properties.title) || [];
    throw new Error(`[${KOKYAKU_VERSION}][ソース読取失敗] file=${spreadsheetId} sheet='${NIPPO_ZENKEN_SHEET_NAME}' available=[${sheetNames.join('|')}] orig=${e.message}`);
  }
  const srcRows = (srcRes.data.values || []).filter(row =>
    row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== '')
  );
  if (srcRows.length === 0) {
    return { count: 0, dateLabel: slashDate };
  }
  console.log(`[顧客更新] コピー元: ${srcRows.length}行`);

  // 3. 利用履歴シートの最初の空白行を特定（A列で判定）
  let targetRes;
  try {
    targetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: KOKYAKU_TARGET_SHEET_ID,
      range: `'${KOKYAKU_TARGET_SHEET_NAME}'!A:A`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
  } catch (e) {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: KOKYAKU_TARGET_SHEET_ID }).catch(() => null);
    const sheetNames = meta?.data?.sheets?.map(s => s.properties.title) || [];
    throw new Error(`[${KOKYAKU_VERSION}][ターゲット読取失敗] file=${KOKYAKU_TARGET_SHEET_ID} sheet='${KOKYAKU_TARGET_SHEET_NAME}' available=[${sheetNames.join('|')}] orig=${e.message}`);
  }
  const targetACol = targetRes.data.values || [];
  let firstEmptyRow = targetACol.length + 1;
  // A列の末尾から空白を遡って最初の空白行を特定
  for (let i = targetACol.length - 1; i >= 0; i--) {
    if (targetACol[i] && targetACol[i][0] && String(targetACol[i][0]).trim() !== '') {
      firstEmptyRow = i + 2; // 0-indexed → 1-indexed + 次の行
      break;
    }
  }
  console.log(`[顧客更新] 貼り付け開始行: ${firstEmptyRow}`);

  // 4. A列に日付、B:J列にデータを書き込み
  const dateValues = srcRows.map(() => [slashDate]);
  const batchData = [
    {
      range: `'${KOKYAKU_TARGET_SHEET_NAME}'!A${firstEmptyRow}`,
      values: dateValues,
    },
    {
      range: `'${KOKYAKU_TARGET_SHEET_NAME}'!B${firstEmptyRow}`,
      values: srcRows,
    },
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: KOKYAKU_TARGET_SHEET_ID,
    requestBody: { valueInputOption: 'USER_ENTERED', data: batchData },
  });
  console.log(`[顧客更新] 書き込み完了: ${srcRows.length}行`);

  // 5. 利用履歴キャッシュをクリア（電話番号検索が最新データを使えるように）
  const defaultEntry = rirekiCacheMap.get(KOKYAKU_TARGET_SHEET_ID);
  if (defaultEntry) defaultEntry.expiresAt = 0;

  return { count: srcRows.length, dateLabel: slashDate };
}

// ── 利用履歴シート ────────────────────────────────────────
const RIREKI_SHEET_ID   = KOKYAKU_TARGET_SHEET_ID;
const RIREKI_SHEET_NAME = KOKYAKU_TARGET_SHEET_NAME;

// 利用履歴専用のSA認証（エスタマOAuthと枠を分けて429を防止）
function createSAAuthClient() {
  const saKeyPath = path.join(process.env.HOME, '.config/chatbot-service-account.json');
  if (fs.existsSync(saKeyPath)) {
    const key = JSON.parse(fs.readFileSync(saKeyPath));
    const auth = new google.auth.GoogleAuth({
      credentials: key,
      scopes: [
        'https://www.googleapis.com/auth/spreadsheets.readonly',
        'https://www.googleapis.com/auth/drive.readonly',
      ],
    });
    return auth;
  }
  // SAキーが無い場合はOAuthにフォールバック
  return createAuthClient();
}

// 利用履歴キャッシュ（spreadsheetId別、A:J全体を保持、2時間）
// key: spreadsheetId, value: { allRows, phoneIndex, expiresAt }
const rirekiCacheMap = new Map();
const rirekiCacheLoadingMap = new Map();

function getRirekiCacheEntry(ssid) {
  let entry = rirekiCacheMap.get(ssid);
  if (!entry) {
    entry = { allRows: null, phoneIndex: null, expiresAt: 0 };
    rirekiCacheMap.set(ssid, entry);
  }
  return entry;
}

// 電話番号を正規化（ハイフン・スペース・全角数字除去）
function normalizePhone(tel) {
  if (!tel) return '';
  const s = tel.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  return s.replace(/[^0-9]/g, '');
}

// A:J列全体を一括ロードしてキャッシュ（2時間保持・リトライ付き）
const RIREKI_CACHE_TTL = 2 * 60 * 60 * 1000; // 2時間

async function buildRirekiCache(ssid = RIREKI_SHEET_ID) {
  const entry = getRirekiCacheEntry(ssid);
  if (entry.allRows && Date.now() < entry.expiresAt) return;
  // 別リクエストが既にリロード中なら既存キャッシュで応答
  if (rirekiCacheLoadingMap.get(ssid)) {
    if (entry.allRows) return;
    // 初回ロード中は少し待つ
    await new Promise(r => setTimeout(r, 5000));
    if (entry.allRows) return;
  }
  rirekiCacheLoadingMap.set(ssid, true);
  try {
    await _loadRirekiWithRetry(ssid);
  } finally {
    rirekiCacheLoadingMap.set(ssid, false);
  }
}

async function _loadRirekiWithRetry(ssid, maxRetries = 3) {
  const entry = getRirekiCacheEntry(ssid);
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`[利用履歴] A:J列一括ロード開始 ssid=${ssid.slice(0,8)}…（${attempt}/${maxRetries}・SA認証）`);
      const auth = createSAAuthClient();
      const sheets = google.sheets({ version: 'v4', auth });
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: ssid,
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
      entry.allRows = rows;
      entry.phoneIndex = phoneIndex;
      entry.expiresAt = Date.now() + RIREKI_CACHE_TTL;
      console.log(`[利用履歴] キャッシュ完了 ssid=${ssid.slice(0,8)}…: ${rows.length}行, ${Object.keys(phoneIndex).length}件の電話番号`);
      return;
    } catch (err) {
      const is429 = err.message && err.message.includes('429');
      if (is429 && attempt < maxRetries) {
        const wait = attempt * 10 * 1000; // 10秒→20秒→30秒
        console.log(`[利用履歴] 429エラー → ${wait / 1000}秒待機してリトライ`);
        await new Promise(r => setTimeout(r, wait));
        continue;
      }
      // リトライ失敗 or 429以外のエラー
      if (entry.allRows) {
        console.log(`[利用履歴] キャッシュ更新失敗（${err.message}）→ 既存データで継続`);
        entry.expiresAt = Date.now() + 10 * 60 * 1000; // 10分後に再試行
        return;
      }
      throw err;
    }
  }
}

// 電話番号で利用履歴を検索（J列に部分一致＝シート側CONTAINSと同じ挙動）
// nameKeyword 指定時は C列(源氏名) にも部分一致するものだけに絞り込み（AND）
// ssid 指定時はその顧客管理SSを検索（店舗別 #T-XXX コマンド用）
async function searchRirekiByPhone(phone, allRecords = false, nameKeyword = '', ssid = RIREKI_SHEET_ID) {
  const normalized = normalizePhone(phone);
  if (normalized.length < 9) return null; // 最低9桁
  await buildRirekiCache(ssid);
  const entry = getRirekiCacheEntry(ssid);
  const rows = entry.allRows || [];
  const kw = nameKeyword ? String(nameKeyword).trim() : '';
  const targets = [];
  // 末尾(=新しい行)から走査して降順に揃える
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

// 場所から部屋番号（末尾の半角/全角スペース＋3〜4桁数字）を除去
function removeRoomNumber(place) {
  return place.replace(/[\s　]+[\d０-９]{3,4}$/, '').trim();
}

// 利用履歴の検索結果をテキストに整形（1件2行・A〜G列表示）
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

// 電話番号パターン（0から始まる10〜11桁、ハイフンあり/なし対応）
// 対応パターン:
//   ・「09012345678」                  → 番号のみ
//   ・「09012345678 谷村」              → 番号 + 名前（C列で部分一致AND）
//   ・「09012345678 全件」「全件 0901…」→ 全件モード（既存互換）
function parsePhoneCommand(text) {
  const normalized = text.replace(/[　]/g, ' ').trim();
  const tokens = normalized.split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return null;

  // 「全件 番号」形式
  if (tokens[0] === '全件' && tokens.length === 2) {
    const phone = tokens[1];
    const d = normalizePhone(phone);
    if (/^[0-9０-９\-ー－]+$/.test(phone) && d.length >= 9 && d.length <= 11 && d.startsWith('0')) {
      return { phone, allRecords: true, nameKeyword: '' };
    }
  }

  // 1個目が電話番号（半角/全角/ハイフン許容、9〜11桁）
  const head = tokens[0];
  const headDigits = normalizePhone(head);
  if (/^[0-9０-９\-ー－]+$/.test(head) && headDigits.length >= 9 && headDigits.length <= 11 && headDigits.startsWith('0')) {
    if (tokens.length === 1) {
      return { phone: head, allRecords: true, nameKeyword: '' };
    }
    if (tokens.length === 2) {
      if (tokens[1] === '全件') return { phone: head, allRecords: true, nameKeyword: '' };
      return { phone: head, allRecords: true, nameKeyword: tokens[1] };
    }
  }
  return null;
}

// 後方互換: 電話番号のみかチェック
function isPhoneNumber(text) {
  return parsePhoneCommand(text) !== null;
}

// 「T-001 09012345678」「T-001 09012345678 谷村」「T-001 09012345678 全件」「T-001 全件 09012345678」を解析
// 戻り値: { storeId, phone, allRecords, nameKeyword } / null
function parseStorePhoneCommand(text) {
  // 全角英数・全角ハイフンを半角化してから判定
  let normalized = String(text || '')
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[－ー―−–—]/g, '-')
    .replace(/[　]/g, ' ')
    .trim();
  // T-001 / T001 / T-1043 などを許容
  const m = normalized.match(/^([Tt]-?\d{3,5})\s+(.+)$/);
  if (!m) return null;
  let storeId = m[1].toUpperCase();
  if (!storeId.includes('-')) storeId = storeId.replace(/^T(\d+)/, 'T-$1');
  const rest = m[2].trim();
  const phoneCmd = parsePhoneCommand(rest);
  if (!phoneCmd) return null;
  return { storeId, ...phoneCmd };
}

// 店舗フォルダ内から「■1.顧客管理」スプレッドシートを検索（IDキャッシュ付き）
const kokyakuKanriIdCache = new Map(); // storeId → spreadsheetId（24時間）
const KOKYAKU_KANRI_NAME = '■1.顧客管理';
const KOKYAKU_KANRI_ID_TTL = 24 * 60 * 60 * 1000;

async function findKokyakuKanriSpreadsheetId(storeId, folderId) {
  const cached = kokyakuKanriIdCache.get(storeId);
  if (cached && Date.now() < cached.expiresAt) return cached.id;
  const auth = createSAAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  // 名前先頭一致で検索（末尾スペース等の揺れに対応）
  const res = await drive.files.list({
    q: `'${folderId}' in parents and trashed=false and mimeType='application/vnd.google-apps.spreadsheet' and name contains '${KOKYAKU_KANRI_NAME}'`,
    fields: 'files(id,name)',
    pageSize: 10,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });
  const files = res.data.files || [];
  if (files.length === 0) {
    throw new Error(`${storeId} のフォルダに「${KOKYAKU_KANRI_NAME}」が見つかりません`);
  }
  // 名前先頭一致で最も短い（末尾装飾なし）を採用
  files.sort((a, b) => a.name.length - b.name.length);
  const id = files[0].id;
  kokyakuKanriIdCache.set(storeId, { id, expiresAt: Date.now() + KOKYAKU_KANRI_ID_TTL });
  return id;
}

// バックグラウンドでプリロード（起動時に呼び出す）
function preloadPhoneIndex() {
  setImmediate(() => buildRirekiCache().catch(e => console.error('[利用履歴] プリロードエラー:', e.message)));
}

// ── LINE イベントハンドラ（200を即返してバックグラウンド処理） ──
async function handleLineEvent(event) {
  const client = createLineClient();

  if (event.type !== 'message' || event.message.type !== 'text') return;

  const text = event.message.text.trim();
  const userId = event.source?.userId;

  // #whoami: 自分のLINE userIdを返信（1対1/グループ共通・許可リスト登録用）
  if (/^#whoami$/i.test(text)) {
    const replyToken = event.replyToken;
    const target = event.source?.groupId || userId;
    console.log(`[LINE] #whoami リクエスト userId=${userId || '(なし)'} target=${target}`);
    setImmediate(async () => {
      const msg = userId
        ? `🆔 あなたのLINE userId:\n${userId}\n\nオーナーにこのIDを伝えて「📁 店舗フォルダ」シートのD列に登録してもらってください。登録後は店舗名と日付だけで日報入力できます。`
        : `❌ userIdを取得できませんでした。botを友達追加してから再度お試しください。`;
      await client.replyMessage({ replyToken, messages: [{ type: 'text', text: msg }] }).catch(() => {});
    });
    return;
  }

  // [C-024 L1] 全コマンド認証: 許可LINE userIdのみ受付（#whoami は上で先に処理済み）
  let allowedStores = [];
  if (userId) {
    try {
      allowedStores = await getStoresByUserId(userId);
    } catch (err) {
      console.error('[認証] エラー:', err.message);
    }
  }
  if (allowedStores.length === 0) {
    console.log(`[認証] 拒否 userId=${userId || '(なし)'} text="${text.slice(0,40)}"`);
    await client.replyMessage({
      replyToken: event.replyToken,
      messages: [{ type: 'text', text: '❌ このLINEは登録されていないため利用できません。\n\n#whoami と送信して自分のLINE userIdを確認し、オーナーに「📁 店舗フォルダ」シートD列への登録を依頼してください。' }],
    }).catch(() => {});
    return;
  }
  console.log(`[認証] 許可 userId=${userId} 店舗=${allowedStores.map(s => s.storeId+'/'+s.storeName).join(',')}`);

  // 「T-001 09012345678 [名前/全件]」形式: 店舗別 顧客検索（1対1/グループ両対応・許可userIdチェック付き）
  const storePhoneCmd = parseStorePhoneCommand(text);
  if (storePhoneCmd) {
    const replyToken = event.replyToken;
    const sourceType = event.source?.type || 'user';
    console.log(`[LINE] 店舗別顧客検索: ${storePhoneCmd.storeId} / ${storePhoneCmd.phone} 名前=${storePhoneCmd.nameKeyword || '(なし)'} source=${sourceType}`);
    setImmediate(async () => {
      try {
        // 1. 店舗ID + 発言者userId で許可チェック → フォルダID取得（グループでも発言者単位で判定）
        const folderInfo = await getFolderByStoreId(storePhoneCmd.storeId, userId);
        // 2. フォルダ内から「■1.顧客管理」を検索
        const kokyakuSsid = await findKokyakuKanriSpreadsheetId(folderInfo.storeId, folderInfo.folderId);
        // 3. 利用履歴を検索（既存ロジックを spreadsheetId 指定で再利用）
        const result = await searchRirekiByPhone(storePhoneCmd.phone, storePhoneCmd.allRecords, storePhoneCmd.nameKeyword, kokyakuSsid);
        console.log(`[利用履歴] ${folderInfo.storeId}/${folderInfo.storeName} 検索結果: ${result ? result.total + '件' : '該当なし'}`);
        const headerKey = storePhoneCmd.nameKeyword ? `${storePhoneCmd.phone} / ${storePhoneCmd.nameKeyword}` : storePhoneCmd.phone;
        const storePrefix = `📍 ${folderInfo.storeId} ${folderInfo.storeName}\n`;
        const msg = (result && result.total > 0)
          ? storePrefix + formatRirekiResult(storePhoneCmd.phone, result, storePhoneCmd.allRecords, storePhoneCmd.nameKeyword)
          : `${storePrefix}📋 利用履歴: ${headerKey}\n該当なし`;
        await client.replyMessage({ replyToken, messages: [{ type: 'text', text: msg }] });
      } catch (err) {
        console.error('[店舗別顧客検索] エラー:', err.message);
        await client.replyMessage({
          replyToken,
          messages: [{ type: 'text', text: `❌ ${err.message}` }],
        }).catch(() => {});
      }
    });
    return;
  }

  // 1対1トークのコマンド処理
  if (event.source?.type === 'user' && userId) {
    setImmediate(async () => {
      // 「4月23日 顧客更新」形式: 日報全件→利用履歴コピー
      const kokyaku = parseKokyakuUpdateCommand(text);
      if (kokyaku) {
        const target = userId || process.env.LINE_USER_ID;
        console.log(`[LINE] 顧客更新リクエスト受信: ${kokyaku.dateStr}`);
        setImmediate(async () => {
          try {
            const result = await processKokyakuUpdate(kokyaku.dateStr);
            console.log(`[顧客更新] 完了: ${result.dateLabel} ${result.count}件`);
          } catch (err) {
            console.error('[顧客更新] エラー:', err.message);
            await client.pushMessage({
              to: target,
              messages: [{ type: 'text', text: `顧客更新エラー: ${err.message}` }],
            }).catch(() => {});
          }
        });
        return;
      }

      // #T001 形式・ID省略形式: 日報ダッシュボード入力（userId許可チェック付き）
      if (isNippoInput(text)) {
        const replyToken = event.replyToken;
        console.log(`[LINE] 日報入力リクエスト受信 userId=${userId || '(なし)'}`);
        setImmediate(async () => {
          try {
            const result = await processNippoInput(text, userId);
            console.log(`[日報入力] 完了: ${result.storeName} / ${result.spreadsheetName} / ${result.cellCount}セル`);
          } catch (err) {
            console.error('[日報入力] エラー:', err.message);
            await client.replyMessage({
              replyToken,
              messages: [{ type: 'text', text: `❌ 入力エラー: ${err.message}` }],
            }).catch(() => {});
          }
        });
        return;
      }

      // [C-033] 【明細】形式: 明細書（新フォーマット）への直接書込（T-XXX必須）
      const meisaisyo = parseMeisaisyoRequest(text);
      if (meisaisyo) {
        const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
        console.log(`[LINE] [C-033] 明細書リクエスト受信: ${meisaisyo.storeId} / ${meisaisyo.dateStr} / ${meisaisyo.name} / 交通費=${meisaisyo.transport ?? '指定なし'} / バンス=${meisaisyo.bance ?? 'なし'} / お釣り=${meisaisyo.otsuri ?? 'なし'} → target=${target}`);
        setImmediate(() => processMeisaisyoAndPush(target, client, meisaisyo, userId).catch(err => console.error('[明細書] 未処理エラー:', err.message)));
        return;
      }

      // 【日付、名前】形式: 明細スクショ
      const meisai = parseMeisaiRequest(text);
      if (meisai) {
        const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
        console.log(`[LINE] 明細リクエスト受信: ${meisai.dateStr} / ${meisai.name} / 交通費=${meisai.transportType ?? '指定なし'} → target=${target}`);
        processMeisaiAndPush(target, client, meisai.dateStr, meisai.name, meisai.transportType).catch(err => console.error('[明細] 未処理エラー:', err.message));
        return;
      }

      // 「月報更新」「月報更新 4月7日」「4月7日 月報更新」: 日報→月報同期
      if (text.includes('月報更新')) {
        const target = userId || process.env.LINE_USER_ID;
        setImmediate(async () => {
          try {
            // 日付指定があれば解析（順序不問: 「月報更新 4月7日」「4月7日 月報更新」「月報更新 4/7」）
            let syncDate = null;
            const dateMatch = text.match(/(\d{1,2})[月\/](\d{1,2})日?/);
            if (dateMatch) {
              const jst = new Date(Date.now() + 9 * 3600 * 1000);
              syncDate = new Date(Date.UTC(jst.getUTCFullYear(), parseInt(dateMatch[1]) - 1, parseInt(dateMatch[2])));
            }
            const result = await syncNippoToGeppo(syncDate);
            const fw = result.fuwamoko;
            const combinedSales = result.totalSales + fw.totalSales;
            const combinedHon   = result.total_hon  + fw.total_hon;
            const combinedCount = result.total_count + fw.total_count;
            await client.pushMessage({
              to: target,
              messages: [{ type: 'text', text:
                `✅ 月報更新完了\n` +
                `📅 ${result.label}\n\n` +
                `━━ CREA ━━\n` +
                `💰 総売上: ${result.totalSales.toLocaleString()}円\n` +
                `📊 本数: ${result.total_hon}本（朝${result.am_hon}/昼${result.pm_hon}/夜${result.night_hon}）\n` +
                `👥 出勤: ${result.total_count}人（朝${result.am_count}/昼${result.pm_count}/夜${result.night_count}）\n\n` +
                `━━ ふわもこSPA ━━\n` +
                `💰 総売上: ${fw.totalSales.toLocaleString()}円\n` +
                `📊 本数: ${fw.total_hon}本（朝${fw.am_hon}/昼${fw.pm_hon}/夜${fw.night_hon}）\n` +
                `👥 出勤: ${fw.total_count}人（朝${fw.am_count}/昼${fw.pm_count}/夜${fw.night_count}）\n\n` +
                `━━ 2店舗合算 ━━\n` +
                `💰 総売上: ${combinedSales.toLocaleString()}円\n` +
                `📊 本数: ${combinedHon}本\n` +
                `👥 出勤: ${combinedCount}人`
              }],
            });
            if (result.warnings && result.warnings.length > 0) {
              const grouped = {};
              for (const w of result.warnings) {
                (grouped[w.section] ||= []).push(w.name);
              }
              const lines = Object.entries(grouped)
                .map(([sec, names]) => `・${sec}: ${names.join(', ')}`)
                .join('\n');
              await client.pushMessage({
                to: target,
                messages: [{ type: 'text', text:
                  `⚠️ 空白行不足で書き込めなかった名前があります\n${lines}\n\n該当シートで空白行を追加するか、不要な行を整理してください。`
                }],
              }).catch(() => {});
            }
          } catch (err) {
            console.error('[月報] エラー:', err.message);
            await client.pushMessage({
              to: target,
              messages: [{ type: 'text', text: `❌ 月報更新エラー: ${err.message}` }],
            }).catch(() => {});
          }
        });
        return;
      }

      // 「教えて」: 日報スクショ
      if (text.includes('教えて')) {
        const target = getPushTarget(event);
        console.log(`[LINE] (教えて) 受信 → バックグラウンド処理開始 target=${target}`);
        processAndPush(target, client).catch(err => console.error('[LINE] 未処理エラー:', err.message));
        return;
      }

      // 電話番号: 利用履歴検索（replyMessageで無料送信）
      const phoneCmd = parsePhoneCommand(text);
      if (phoneCmd) {
        const replyToken = event.replyToken;
        console.log(`[LINE] 電話番号検索: ${phoneCmd.phone} 全件=${phoneCmd.allRecords} 名前=${phoneCmd.nameKeyword || '(なし)'}`);
        setImmediate(async () => {
          try {
            const result = await searchRirekiByPhone(phoneCmd.phone, phoneCmd.allRecords, phoneCmd.nameKeyword);
            console.log(`[利用履歴] 検索結果: ${result ? result.total + '件' : '該当なし'}`);
            const headerKey = phoneCmd.nameKeyword ? `${phoneCmd.phone} / ${phoneCmd.nameKeyword}` : phoneCmd.phone;
            const msg = (result && result.total > 0)
              ? formatRirekiResult(phoneCmd.phone, result, phoneCmd.allRecords, phoneCmd.nameKeyword)
              : `📋 利用履歴: ${headerKey}\n該当なし`;
            console.log(`[利用履歴] LINE reply送信（${msg.length}文字）`);
            await client.replyMessage({ replyToken, messages: [{ type: 'text', text: msg }] });
            console.log(`[利用履歴] LINE送信完了`);
          } catch (err) {
            console.error('[利用履歴] エラー詳細:', err.code || err.statusCode || '', err.message);
          }
        });
        return;
      }
    });
    return;
  }

  // グループからのコマンド

  // ════════════════════════════════════════════════════════════════════
  // ▼ [C-036] 出勤表グループ → 新SS自動反映ブロック  ⚠️ 編集禁止 ⚠️
  // ════════════════════════════════════════════════════════════════════
  // このブロックは shift_reflect.js 経由で LINEシフト→月次SSへ自動反映。
  // 削除・改変すると C-036 機能（出勤表LINEのシフト自動反映）が壊れます。
  // 何か変更が必要な場合は shift_reflect.js 側で対応してください。
  // 削除された場合、サーバー上のガーディアン (cron 1分毎) が自動復旧します。
  // 詳細: 共有記憶シート C-036 行 を参照
  //   https://docs.google.com/spreadsheets/d/1-ITHWZsvH9Z6iNTNtEUp4LTM5pWN2VwT6A41SeTBy_c
  // ════════════════════════════════════════════════════════════════════
  const SHIFT_GROUP_ID = 'C4befe4675d94c864734eae6b897f1484';
  if (event.source?.type === 'group' && event.source?.groupId === SHIFT_GROUP_ID) {
    const parts = text.trim().split(/[\s　]+/);
    if (parts.length >= 3 || text === '確認' || text === '出勤確認') {
      setImmediate(async () => {
        try {
          const shiftReflect = require('./shift_reflect');
          const result = await shiftReflect.reflectShiftMessage(text);
          if (result.type === 'ignore') {
            console.log('[Shift] 非シフトメッセージ→スルー');
          } else if (result.type === 'success') {
            console.log(`[Shift] 結果: success ${result.writtenCount} 件 (${result.staffName} / ${result.store})`);
          } else {
            console.log(`[Shift] 結果: ${JSON.stringify(result)}`);
          }
          const reply = shiftReflect.formatReply(result);
          if (reply) {
            await client.pushMessage({
              to: event.source.groupId,
              messages: [{ type: 'text', text: reply }],
            }).catch(() => {});
          }
        } catch (err) {
          console.error('[Shift] エラー:', err.message);
        }
      });
      return;
    }
  }


  // 「4月23日 顧客更新」形式: 日報全件→利用履歴コピー（グループ）
  const kokyakuG = parseKokyakuUpdateCommand(text);
  if (kokyakuG) {
    const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
    console.log(`[LINE] 顧客更新リクエスト受信（グループ）: ${kokyakuG.dateStr}`);
    setImmediate(async () => {
      try {
        const result = await processKokyakuUpdate(kokyakuG.dateStr);
        console.log(`[顧客更新] グループ完了: ${result.dateLabel} ${result.count}件`);
      } catch (err) {
        console.error('[顧客更新] グループ エラー:', err.message);
        await client.pushMessage({
          to: target,
          messages: [{ type: 'text', text: `顧客更新エラー: ${err.message}` }],
        }).catch(() => {});
      }
    });
    return;
  }

  // #T001 形式・ID省略形式: 日報ダッシュボード入力（グループ・userId許可チェック付き）
  if (isNippoInput(text)) {
    const replyToken = event.replyToken;
    console.log(`[LINE] 日報入力リクエスト受信（グループ） userId=${userId || '(なし)'}`);
    setImmediate(async () => {
      try {
        const result = await processNippoInput(text, userId);
        console.log(`[日報入力] グループ完了: ${result.storeName} / ${result.spreadsheetName} / ${result.cellCount}セル`);
      } catch (err) {
        console.error('[日報入力] グループ エラー:', err.message);
        await client.replyMessage({
          replyToken,
          messages: [{ type: 'text', text: `❌ 入力エラー: ${err.message}` }],
        }).catch(() => {});
      }
    });
    return;
  }

  // 「月報更新」「月報更新 4月7日」「4月7日 月報更新」: 日報→月報同期
  if (text.includes('月報更新')) {
    const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
    setImmediate(async () => {
      try {
        let syncDate = null;
        const dateMatch = text.match(/(\d{1,2})[月\/](\d{1,2})日?/);
        if (dateMatch) {
          const jst = new Date(Date.now() + 9 * 3600 * 1000);
          syncDate = new Date(Date.UTC(jst.getUTCFullYear(), parseInt(dateMatch[1]) - 1, parseInt(dateMatch[2])));
        }
        const result = await syncNippoToGeppo(syncDate);
        const fw = result.fuwamoko;
        const combinedSales = result.totalSales + fw.totalSales;
        const combinedHon   = result.total_hon  + fw.total_hon;
        const combinedCount = result.total_count + fw.total_count;
        await client.pushMessage({
          to: target,
          messages: [{ type: 'text', text:
            `✅ 月報更新完了\n` +
            `📅 ${result.label}\n\n` +
            `━━ CREA ━━\n` +
            `💰 総売上: ${result.totalSales.toLocaleString()}円\n` +
            `📊 本数: ${result.total_hon}本（朝${result.am_hon}/昼${result.pm_hon}/夜${result.night_hon}）\n` +
            `👥 出勤: ${result.total_count}人（朝${result.am_count}/昼${result.pm_count}/夜${result.night_count}）\n\n` +
            `━━ ふわもこSPA ━━\n` +
            `💰 総売上: ${fw.totalSales.toLocaleString()}円\n` +
            `📊 本数: ${fw.total_hon}本（朝${fw.am_hon}/昼${fw.pm_hon}/夜${fw.night_hon}）\n` +
            `👥 出勤: ${fw.total_count}人（朝${fw.am_count}/昼${fw.pm_count}/夜${fw.night_count}）\n\n` +
            `━━ 2店舗合算 ━━\n` +
            `💰 総売上: ${combinedSales.toLocaleString()}円\n` +
            `📊 本数: ${combinedHon}本\n` +
            `👥 出勤: ${combinedCount}人`
          }],
        });
        if (result.warnings && result.warnings.length > 0) {
          const grouped = {};
          for (const w of result.warnings) {
            (grouped[w.section] ||= []).push(w.name);
          }
          const lines = Object.entries(grouped)
            .map(([sec, names]) => `・${sec}: ${names.join(', ')}`)
            .join('\n');
          await client.pushMessage({
            to: target,
            messages: [{ type: 'text', text:
              `⚠️ 空白行不足で書き込めなかった名前があります\n${lines}\n\n該当シートで空白行を追加するか、不要な行を整理してください。`
            }],
          }).catch(() => {});
        }
      } catch (err) {
        console.error('[月報] グループ エラー:', err.message);
        const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
        await client.pushMessage({
          to: target,
          messages: [{ type: 'text', text: `❌ 月報更新エラー: ${err.message}` }],
        }).catch(() => {});
      }
    });
    return;
  }

  // [C-033] 【明細】形式: 明細書（新フォーマット）への直接書込（T-XXX必須）
  const meisaisyoG = parseMeisaisyoRequest(text);
  if (meisaisyoG) {
    const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
    console.log(`[LINE] [C-033] 明細書リクエスト受信(G): ${meisaisyoG.storeId} / ${meisaisyoG.dateStr} / ${meisaisyoG.name} / 交通費=${meisaisyoG.transport ?? '指定なし'} / バンス=${meisaisyoG.bance ?? 'なし'} / お釣り=${meisaisyoG.otsuri ?? 'なし'} → target=${target}`);
    setImmediate(() => processMeisaisyoAndPush(target, client, meisaisyoG, userId).catch(err => console.error('[明細書] 未処理エラー:', err.message)));
    return;
  }

  // 【日付、名前】形式: 明細スクショ
  const meisai = parseMeisaiRequest(text);
  if (meisai) {
    const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
    console.log(`[LINE] 明細リクエスト受信: ${meisai.dateStr} / ${meisai.name} / 交通費=${meisai.transportType ?? '指定なし'} → target=${target}`);
    setImmediate(() => processMeisaiAndPush(target, client, meisai.dateStr, meisai.name, meisai.transportType).catch(err => console.error('[明細] 未処理エラー:', err.message)));
    return;
  }

  // 「教えて」: 日報スクショ
  if (text.includes('教えて')) {
    const target = getPushTarget(event);
    console.log(`[LINE] (教えて) 受信 → バックグラウンド処理開始 target=${target}`);
    setImmediate(() => processAndPush(target, client).catch(err => console.error('[LINE] 未処理エラー:', err.message)));
    return;
  }

  // 電話番号: 利用履歴検索（グループ・replyMessageで無料送信）
  const phoneCmdG = parsePhoneCommand(text);
  if (phoneCmdG) {
    const replyTokenG = event.replyToken;
    console.log(`[LINE] 電話番号検索(グループ): ${phoneCmdG.phone} 全件=${phoneCmdG.allRecords} 名前=${phoneCmdG.nameKeyword || '(なし)'}`);
    setImmediate(async () => {
      try {
        const result = await searchRirekiByPhone(phoneCmdG.phone, phoneCmdG.allRecords, phoneCmdG.nameKeyword);
        console.log(`[利用履歴] 検索結果: ${result ? result.total + '件' : '該当なし'}`);
        const headerKeyG = phoneCmdG.nameKeyword ? `${phoneCmdG.phone} / ${phoneCmdG.nameKeyword}` : phoneCmdG.phone;
        const msg = (result && result.total > 0)
          ? formatRirekiResult(phoneCmdG.phone, result, phoneCmdG.allRecords, phoneCmdG.nameKeyword)
          : `📋 利用履歴: ${headerKeyG}\n該当なし`;
        console.log(`[利用履歴] LINE reply送信（${msg.length}文字）`);
        await client.replyMessage({ replyToken: replyTokenG, messages: [{ type: 'text', text: msg }] });
        console.log(`[利用履歴] LINE送信完了`);
      } catch (err) {
        console.error('[利用履歴] エラー詳細:', err.code || err.statusCode || '', err.message);
      }
    });
    return;
  }
}

module.exports = { lineConfig, handleLineEvent, middleware, preloadPhoneIndex, searchRirekiByPhone, formatRirekiResult, NIPPO_FOLDER_ID_DEBUG: NIPPO_FOLDER_ID };
