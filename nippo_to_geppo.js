/**
 * nippo_to_geppo.js
 * 日報（日ごとスプシ）のデータを月報に自動書き込みする
 *
 * 読み取り元: 日報・売上シート
 *   CM3 = 総売上
 *   CO3 = 女の子給料合計
 *   CA4:CK35 = スタッフ別行（CA=出勤時間, CC=本数, CK=待機時間）
 *
 * 書き込み先: 月報・売上シート
 *   列: G=1日, H=2日... Z=20日, AA=21日... AK=31日
 *   行3=総売上, 行5=朝本数, 行6=昼本数, 行7=夜本数, 行8=総本数
 *   行9=朝出勤, 行10=昼出勤, 行11=夜出勤, 行12=合計出勤
 *   行14=待機合計, 行18=女の子給料
 */

require('dotenv').config();
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const GEPPO_SHEET_ID = process.env.GEPPO_SHEET_ID || '1tikBFx4F0mnEx2sxyVSVRb4RROJtNv_auDrmH5orbUw';
const NIPPO_FOLDER_ID = '1isPYyiUqyWXnS1mtpE1_YWJ9QZBTemdJ';

// 朝/昼/夜の出勤時刻境界（例: 12.0 = 12:00）
const AM_END = 12;   // 12時未満 = 朝
const PM_END = 19;   // 12〜18時台 = 昼、19時以降 = 夜

// ── Google認証 ────────────────────────────────────────────
// 書き込みスコープ付きトークンが必要なため、credentialsファイルを優先する
let _authClient = null;
function createAuthClient() {
  if (_authClient) return _authClient;
  let client_id, client_secret, refresh_token;

  const credsPath = path.join(process.env.HOME, '.config/gdrive-server-credentials.json');
  const keysPath  = path.join(process.env.HOME, '.config/gcp-oauth.keys.json');

  if (fs.existsSync(credsPath) && fs.existsSync(keysPath)) {
    // ローカル/サーバー: credentialsファイルから取得（Sheets書き込みスコープ付き）
    const oauthKeys   = JSON.parse(fs.readFileSync(keysPath));
    const credentials = JSON.parse(fs.readFileSync(credsPath));
    const installed   = oauthKeys.installed || oauthKeys.web;
    client_id     = installed.client_id;
    client_secret = installed.client_secret;
    refresh_token = credentials.refresh_token;
  } else {
    // credentialsファイルがない場合のみ環境変数を使用
    client_id     = process.env.GOOGLE_CLIENT_ID;
    client_secret = process.env.GOOGLE_CLIENT_SECRET;
    refresh_token = process.env.GOOGLE_REFRESH_TOKEN;
  }

  const oauth2Client = new google.auth.OAuth2(client_id, client_secret, 'http://localhost');
  oauth2Client.setCredentials({ refresh_token });
  _authClient = oauth2Client;
  return oauth2Client;
}

// ── 日付 → 月報の列文字（G=1日, H=2日...） ────────────────
function dayToCol(day) {
  const idx = 6 + (day - 1); // G=6 (0-indexed: A=0)
  if (idx < 26) return String.fromCharCode(65 + idx);
  return 'A' + String.fromCharCode(65 + idx - 26);
}

// ── 日付文字列から日報ファイルを検索 ──────────────────────
async function findNippoFileId(jstDate) {
  const y = jstDate.getUTCFullYear();
  const m = jstDate.getUTCMonth() + 1;
  const d = jstDate.getUTCDate();
  const dateStr = `${y}年${m}月${d}日`;

  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });

  const res = await drive.files.list({
    q: `'${NIPPO_FOLDER_ID}' in parents and name contains '${dateStr}' and trashed = false`,
    fields: 'files(id, name)',
    orderBy: 'createdTime desc',
    pageSize: 5,
  });

  const files = res.data.files || [];
  console.log(`[月報] 日報検索 "${dateStr}": ${files.map(f => f.name).join(', ') || 'なし'}`);

  // 「2026年4月7日」に完全一致するものを優先（スペースや別店舗名を除外）
  const exact = files.find(f => f.name.trim() === dateStr);
  return (exact || files[0])?.id || null;
}

// ── 日報ファイルからデータ読み取り ────────────────────────
async function readNippoData(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // CM3（CREA総売上）+ CO3（女子給与合計）
  const r1 = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'売上'!CL3:CQ3",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const totRow = r1.data.values?.[0] || [];
  // CL=0 ('C'), CM=1 (総売上), CO=3 (女子給), CQ=5 (店取り分)
  const totalSales  = Number(totRow[1]) || 0;
  const staffSalary = Number(totRow[3]) || 0;

  // CA4:CK35（CREA + ふわもこSPA 全スタッフ行）
  // CA=出勤, CB=退勤, CC=本数, CD=名前, CK=待機時間
  const r2 = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'売上'!CA4:CK35",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r2.data.values || [];

  let am_hon = 0, pm_hon = 0, night_hon = 0;
  let am_count = 0, pm_count = 0, night_count = 0;
  let total_stay = 0;

  for (const row of rows) {
    const shukkin = Number(row[0]); // CA: 出勤時刻（例: 11, 15.5, 23.5）
    const honsu   = Number(row[2]); // CC: 本数
    const stay    = Number(row[10]);// CK: 待機時間

    // 出勤時刻が数値でない行（ヘッダー行など）はスキップ
    if (!row[0] || isNaN(shukkin) || shukkin === 0) continue;
    if (isNaN(honsu)) continue;

    total_stay += isNaN(stay) ? 0 : stay;

    if (shukkin < AM_END) {
      am_count++;
      am_hon += honsu;
    } else if (shukkin < PM_END) {
      pm_count++;
      pm_hon += honsu;
    } else {
      night_count++;
      night_hon += honsu;
    }
  }

  const total_hon   = am_hon + pm_hon + night_hon;
  const total_count = am_count + pm_count + night_count;

  return { totalSales, staffSalary, am_hon, pm_hon, night_hon, total_hon, am_count, pm_count, night_count, total_count, total_stay };
}

// ── 月報シートへ書き込み ───────────────────────────────────
async function writeToGeppo(jstDate, data) {
  const day = jstDate.getUTCDate();
  const col = dayToCol(day);

  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const updates = [
    { range: `'売上'!${col}3`,  values: [[data.totalSales]]  },  // 総売上
    { range: `'売上'!${col}5`,  values: [[data.am_hon]]      },  // 朝本数
    { range: `'売上'!${col}6`,  values: [[data.pm_hon]]      },  // 昼本数
    { range: `'売上'!${col}7`,  values: [[data.night_hon]]   },  // 夜本数
    { range: `'売上'!${col}8`,  values: [[data.total_hon]]   },  // 総本数
    { range: `'売上'!${col}9`,  values: [[data.am_count]]    },  // 朝出勤
    { range: `'売上'!${col}10`, values: [[data.pm_count]]    },  // 昼出勤
    { range: `'売上'!${col}11`, values: [[data.night_count]] },  // 夜出勤
    { range: `'売上'!${col}12`, values: [[data.total_count]] },  // 合計出勤
    { range: `'売上'!${col}14`, values: [[data.total_stay]]  },  // 待機合計
    { range: `'売上'!${col}18`, values: [[data.staffSalary]] },  // 女の子給料
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: GEPPO_SHEET_ID,
    requestBody: { valueInputOption: 'RAW', data: updates },
  });

  console.log(`[月報] ${day}日(${col}列) 書き込み完了:`, data);
  return { day, col, ...data };
}

// ── メイン関数: 指定日の日報→月報同期 ────────────────────
// jstDate: JST基準のDateオブジェクト（省略時は昨日）
async function syncNippoToGeppo(jstDate) {
  if (!jstDate) {
    // デフォルト: JST昨日（深夜4時実行なら前営業日）
    const now = new Date();
    const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    jst.setUTCDate(jst.getUTCDate() - 1);
    jstDate = jst;
  }

  const y = jstDate.getUTCFullYear();
  const m = jstDate.getUTCMonth() + 1;
  const d = jstDate.getUTCDate();
  const label = `${y}年${m}月${d}日`;

  console.log(`[月報] ${label} の同期開始`);

  const fileId = await findNippoFileId(jstDate);
  if (!fileId) throw new Error(`日報ファイルが見つかりません: ${label}`);

  const data = await readNippoData(fileId);
  const result = await writeToGeppo(jstDate, data);

  return { label, ...result };
}

module.exports = { syncNippoToGeppo, dayToCol, findNippoFileId };
