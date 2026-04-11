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

const GEPPO_SHEET_ID        = process.env.GEPPO_SHEET_ID        || '1L5a0SeqSckZARYq3rZBVpBwDBPL4GyliEqA-TIDzW7Y'; // CREA月報（フォールバック）
const FUWAMOKO_GEPPO_ID     = process.env.FUWAMOKO_GEPPO_ID     || '1wYQ5YYU9zSbGZ7VDnPp89mTMX_JGIUdNNNFBVgBneQ4'; // ふわもこSPA月報（フォールバック）
const NIPPO_FOLDER_ID = '1isPYyiUqyWXnS1mtpE1_YWJ9QZBTemdJ';

// ── 月報シートIDを動的に取得（フォルダ内の「C売上YYYY-M月」を検索） ────
// prefix: 'C'=CREA, 'F'=ふわもこSPA
async function findGeppoSheetId(year, month, prefix) {
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const name = `${prefix}売上${year}-${month}月`;
  const res = await drive.files.list({
    q: `'${NIPPO_FOLDER_ID}' in parents and name='${name}' and trashed=false`,
    fields: 'files(id, name, mimeType, shortcutDetails)',
  });
  const files = res.data.files || [];
  if (files.length === 0) {
    console.log(`[月報] ${name} が見つかりません。フォールバックIDを使用します。`);
    return null;
  }
  const file = files[0];
  // ショートカットの場合はターゲットIDを返す
  if (file.mimeType === 'application/vnd.google-apps.shortcut' && file.shortcutDetails) {
    console.log(`[月報] ${name} → ショートカット先: ${file.shortcutDetails.targetId}`);
    return file.shortcutDetails.targetId;
  }
  console.log(`[月報] ${name} → ${file.id}`);
  return file.id;
}

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

// ── 日付 → 月報売上シートの列文字（G=1日, H=2日...） ────────
function dayToCol(day) {
  const idx = 6 + (day - 1); // G=6 (0-indexed: A=0)
  if (idx < 26) return String.fromCharCode(65 + idx);
  return 'A' + String.fromCharCode(65 + idx - 26);
}

// ── 日付 → 回収シートの列文字（E=1日, F=2日...） ────────────
function dayToKaishuuCol(day) {
  const idx = 4 + (day - 1); // E=4 (0-indexed: A=0)
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
    q: `'${NIPPO_FOLDER_ID}' in parents and name = '${dateStr}' and trashed = false`,
    fields: 'files(id, name)',
    pageSize: 5,
  });

  const files = res.data.files || [];
  console.log(`[月報] 日報検索 "${dateStr}": ${files.map(f => f.name).join(', ') || 'なし'}`);

  // 完全一致のみ返す（部分一致・別店舗は除外）
  const exact = files.find(f => f.name.trim() === dateStr);
  return exact?.id || null;
}

// CA/CB（出勤・退勤時刻）から朝/昼/夜の出勤カウントを計算
// シフトが複数時間帯をまたぐ場合は各帯でカウント
function calcShukkinBySlot(caRows, startRow, endRow) {
  let am = 0, pm = 0, night = 0;
  const overlaps = (shiftStart, shiftEnd, ps, pe) => shiftStart < pe && shiftEnd > ps;

  for (let i = startRow - 1; i <= endRow - 1; i++) { // 0-indexed (行3=index0 for CA3:CB35)
    const row = caRows[i];
    if (!row) continue;
    const ca = Number(row[0]);
    const cb = Number(row[1]);
    if (isNaN(ca) || ca === 0 || isNaN(cb) || cb === 0) continue;
    if (typeof row[0] === 'string') continue; // ヘッダー行スキップ

    if (overlaps(ca, cb, 9,  15)) am++;
    if (overlaps(ca, cb, 15, 21)) pm++;
    if (overlaps(ca, cb, 21, 30)) night++;
  }
  return { am, pm, night };
}

// BD列の小数時刻 → 時間（h）に変換（翌日分は+24）
function bdToHour(bd) {
  let h = bd * 24;
  if (h < 9) h += 24; // 深夜0〜8時台 → 24〜32時扱い
  return h;
}

// 時間（h）→ 朝/昼/夜 分類
function timeSlot(h) {
  if (h >= 9  && h < 15) return 'am';
  if (h >= 15 && h < 21) return 'pm';
  if (h >= 21 && h <= 30) return 'night';
  return null;
}

// BD列から朝/昼/夜本数を集計（startRow〜endRow: 1-indexed、BD列はindex0）
function calcHonsuByBd(bdRows, startRow, endRow) {
  let am = 0, pm = 0, night = 0;
  for (let i = startRow - 2; i <= endRow - 2; i++) { // BD2=index0
    const row = bdRows[i];
    if (!row) continue;
    const bd = Number(row[0]);
    if (isNaN(bd) || bd <= 0 || bd > 10) continue; // 異常値スキップ
    const slot = timeSlot(bdToHour(bd));
    if (slot === 'am')    am++;
    else if (slot === 'pm')    pm++;
    else if (slot === 'night') night++;
  }
  return { am, pm, night };
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
  const totalSales  = Number(totRow[1]) || 0;
  const staffSalary = Number(totRow[3]) || 0;

  // CA4:CK21（CREAスタッフ行）- 出勤数・待機合計用
  const r2 = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'売上'!CA4:CK21",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r2.data.values || [];

  let total_stay = 0;
  let total_count = 0;
  for (const row of rows) {
    if (!row[0] || typeof row[0] === 'string') continue;
    const ca = Number(row[0]);
    if (isNaN(ca) || ca === 0) continue;
    total_count++; // 行にデータが入っている人数（実出勤人数）
    const stay = Number(row[10]);
    total_stay += isNaN(stay) ? 0 : stay;
  }

  // CA/CB列でシフト重複ありの朝/昼/夜出勤カウント（1スタッフが複数時間帯をまたぐ場合は複数カウント）
  const { am: am_count, pm: pm_count, night: night_count } = calcShukkinBySlot(rows, 1, rows.length);

  // BD列（CREA: 行3〜46）で朝/昼/夜本数を集計
  const rBd = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'売上'!BD2:BD46",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const { am: am_hon, pm: pm_hon, night: night_hon } = calcHonsuByBd(rBd.data.values || [], 3, 46);
  const total_hon = am_hon + pm_hon + night_hon;

  return { totalSales, staffSalary, am_hon, pm_hon, night_hon, total_hon, am_count, pm_count, night_count, total_count, total_stay };
}

// ── ふわもこSPAデータ読み取り ─────────────────────────────
async function readFuwamokoData(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // CM4（ふわもこ総売上）+ CO4（女子給与）
  const r1 = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'売上'!CL4:CQ4",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const totRow = r1.data.values?.[0] || [];
  const totalSales  = Number(totRow[1]) || 0;  // CM4
  const staffSalary = Number(totRow[3]) || 0;  // CO4

  // CA23:CK32（ふわもこSPAスタッフ行）- 出勤数・待機合計用
  const r2 = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'売上'!CA23:CK32",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r2.data.values || [];

  let total_stay = 0;
  let total_count = 0;
  for (const row of rows) {
    if (!row[0] || typeof row[0] === 'string') continue;
    const ca = Number(row[0]);
    if (isNaN(ca) || ca === 0) continue;
    total_count++; // 行にデータが入っている人数（実出勤人数）
    const stay = Number(row[10]);
    total_stay += isNaN(stay) ? 0 : stay;
  }

  // CA/CB列でシフト重複ありの朝/昼/夜出勤カウント
  const { am: am_count, pm: pm_count, night: night_count } = calcShukkinBySlot(rows, 1, rows.length);

  // BD列（ふわもこ: 行48〜64）で朝/昼/夜本数を集計
  const rBd = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'売上'!BD2:BD64",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const { am: am_hon, pm: pm_hon, night: night_hon } = calcHonsuByBd(rBd.data.values || [], 48, 64);
  const total_hon = am_hon + pm_hon + night_hon;

  return { totalSales, staffSalary, am_hon, pm_hon, night_hon, total_hon, am_count, pm_count, night_count, total_count, total_stay };
}

// ── 月報シートへの新規名前自動登録 ────────────────────────────
// 日報に現れた新しい名前をSBシート(A列)・回収シート(D列)に上から順に追加
// dataStartRow: 名前を書き始める最初の行（SB=9, 回収=2）
async function autoRegisterNames(names, sheetId, tabName, colLetter, dataStartRow, label) {
  if (!names || names.length === 0) return;

  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `'${tabName}'!${colLetter}1:${colLetter}300`,
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = res.data.values || [];
  const existing = rows.map(r => normalizeName((r[0] || '').toString()));

  // まだ登録されていない名前だけ絞り込む
  const toAdd = names.filter(name => {
    const norm = normalizeName(name);
    return norm && !existing.some(e => e === norm);
  });

  if (toAdd.length === 0) {
    console.log(`[名前登録/${label}] 新規なし`);
    return;
  }

  // dataStartRow から下に向かって最初の空白行を探す
  let nextRow = dataStartRow;
  for (let i = dataStartRow - 1; i < rows.length; i++) {
    const val = normalizeName((rows[i]?.[0] || '').toString());
    if (!val) { nextRow = i + 1; break; }  // 空白行を発見
    nextRow = i + 2; // まだ埋まっている → 次の行へ
  }

  const updates = toAdd.map((name, i) => ({
    range: `'${tabName}'!${colLetter}${nextRow + i}`,
    values: [[name]],
  }));

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: { valueInputOption: 'RAW', data: updates },
  });

  console.log(`[名前登録/${label}] ${toAdd.length}件追加(${nextRow}行〜):`, toAdd);
  return toAdd;
}

// ── 月報シートへ書き込み（sheetId指定可） ────────────────────
async function writeToGeppo(jstDate, data, sheetId, label) {
  sheetId = sheetId || GEPPO_SHEET_ID;
  label   = label   || 'CREA';
  const day = jstDate.getUTCDate();
  const col = dayToCol(day);

  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const updates = [
    { range: `'売上'!${col}3`,  values: [[data.totalSales]]  },  // 総売上
    { range: `'売上'!${col}5`,  values: [[data.am_hon]]      },  // 朝本数
    { range: `'売上'!${col}6`,  values: [[data.pm_hon]]      },  // 昼本数
    { range: `'売上'!${col}7`,  values: [[data.night_hon]]   },  // 夜本数
    { range: `'売上'!${col}8`,  values: [[data.total_hon]]    },  // 総本数
    { range: `'売上'!${col}9`,  values: [[data.am_count]]    },  // 朝出勤
    { range: `'売上'!${col}10`, values: [[data.pm_count]]    },  // 昼出勤
    { range: `'売上'!${col}11`, values: [[data.night_count]] },  // 夜出勤
    { range: `'売上'!${col}12`, values: [[data.total_count]] },  // 合計出勤
    { range: `'売上'!${col}14`, values: [[data.total_stay]]  },  // 待機合計
    { range: `'売上'!${col}18`, values: [[data.staffSalary]] },  // 女の子給料
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: { valueInputOption: 'RAW', data: updates },
  });

  console.log(`[月報/${label}] ${day}日(${col}列) 書き込み完了:`, data);
  return { day, col, ...data };
}

// 漢字を含む名前 = CREAスタッフ判定（ひらがなのみ = ふわもこSPA）
function isCrea(name) {
  return /[\u4e00-\u9fff\u3400-\u4dbf]/.test(name);
}

// 「【F】山本」→「山本」、「立花 特」→「立花」のように括弧・スペース以降を除去
function normalizeName(str) {
  return str
    .replace(/【[^】]*】/g, '')  // 【...】除去
    .replace(/\[[^\]]*\]/g, '') // [...] 除去
    .split(/[\s　]/)[0]          // スペース（全角含む）以降を除去
    .trim();
}

// ── 日報・手持ちシートから経費（交通費）読み取り ────────────
// 戻り値: { crea: [...], fuwamoko: [...] }
async function readKeihi(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'手持ち'!K42:N80",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r.data.values || [];

  // 行42はヘッダー（名前/給料/交通費）なのでスキップ、43行目以降がデータ
  const crea = [], fuwamoko = [];
  for (const row of rows.slice(1)) {
    const name  = row[0];
    const kotsu = Number(row[3]);
    if (!name || typeof name !== 'string' || name.trim() === '') continue;
    if (!(kotsu > 0)) continue;  // 交通費0円 → スキップ
    const item = { name: name.trim(), kotsu };
    if (isCrea(name)) {
      crea.push(item);      // 漢字名 = CREA
    } else {
      fuwamoko.push(item);  // ひらがな名 = ふわもこSPA
    }
  }

  return { crea, fuwamoko };
}

// ── 月報・経費シートへ書き込み（指定シートID） ───────────
async function writeKeihi(jstDate, keihi, sheetId, label) {
  if (keihi.length === 0) {
    console.log(`[経費/${label}] 書き込み対象なし`);
    return [];
  }

  const day = jstDate.getUTCDate();
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // 経費シートのC・D・H列を読んで既存データと空き行を確認
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: "'経費'!C1:H300",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const existingRows = res.data.values || [];

  // 既存データ（日・名前・金額）をセットに入れて重複チェック用に使う
  const existingSet = new Set();
  let startRow = -1;
  for (let i = 3; i < existingRows.length; i++) {
    const row = existingRows[i];
    const c = row?.[0], d = row?.[1], h = row?.[5];
    if (c === '' || c === undefined || c === null) {
      if (startRow === -1) startRow = i + 1; // 最初の空き行（1-indexed）
    } else {
      existingSet.add(`${c}|${d}|${h}`);
    }
  }
  if (startRow === -1) startRow = existingRows.length + 1;

  // 重複を除外
  const newKeihi = keihi.filter(item => {
    const key = `${day}|${item.name}|${item.kotsu}`;
    if (existingSet.has(key)) {
      console.log(`[経費/${label}] スキップ(重複): ${day}日 ${item.name} ${item.kotsu}円`);
      return false;
    }
    return true;
  });

  if (newKeihi.length === 0) {
    console.log(`[経費/${label}] 書き込み対象なし（全件重複）`);
    return [];
  }

  const updates = newKeihi.map((item, idx) => {
    const row = startRow + idx;
    return [
      { range: `'経費'!C${row}`, values: [[day]]        },  // 日
      { range: `'経費'!D${row}`, values: [[item.name]]   },  // 名前
      { range: `'経費'!H${row}`, values: [[item.kotsu]]  },  // 交通費
    ];
  }).flat();

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: { valueInputOption: 'RAW', data: updates },
  });

  console.log(`[経費/${label}] ${day}日 ${newKeihi.length}件 書き込み完了(${startRow}行目〜):`, newKeihi);
  return newKeihi;
}

// ── 日報・手持ちシートから回収データ読み取り ─────────────────
// 戻り値: { crea: [{name, amount}], fuwamoko: [{name, amount}] }
async function readKaishuu(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'手持ち'!K42:Q80",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r.data.values || [];

  const crea = [], fuwamoko = [];
  for (const row of rows.slice(1)) {
    const name   = row[0];
    const o      = Number(row[4]); // O列（前回）
    const p      = Number(row[5]); // P列（次回）
    if (!name || typeof name !== 'string' || name.trim() === '') continue;
    if (isNaN(o) && isNaN(p)) continue;
    const amount = (isNaN(o) ? 0 : o) + (isNaN(p) ? 0 : p);
    if (amount === 0) continue;
    const item = { name: name.trim(), amount };
    if (isCrea(name)) {
      crea.push(item);
    } else {
      fuwamoko.push(item);
    }
  }
  return { crea, fuwamoko };
}

// ── 月報・回収シートへ書き込み ────────────────────────────
// 回収シートのD列で名前を検索し、日付列（E=1日...）にamountを書き込む
async function writeKaishuu(jstDate, kaishuu, sheetId, label) {
  if (kaishuu.length === 0) {
    console.log(`[回収/${label}] 書き込み対象なし`);
    return [];
  }

  const day = jstDate.getUTCDate();
  const col = dayToKaishuuCol(day);
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // D列の名前リストを取得
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: "'回収'!D1:D200",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const dCol = (res.data.values || []).map(r => normalizeName((r[0] || '').toString()));

  const updates = [];
  const notFound = [];

  for (const item of kaishuu) {
    const rowIdx = dCol.findIndex((name, i) => i > 0 && name === item.name);
    if (rowIdx === -1) {
      notFound.push(item.name);
      continue;
    }
    const row = rowIdx + 1; // 1-indexed
    updates.push({ range: `'回収'!${col}${row}`, values: [[-item.amount]] }); // 符号反転（-→+、+→-）
  }

  if (updates.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: { valueInputOption: 'RAW', data: updates },
    });
  }

  if (notFound.length > 0) console.log(`[回収/${label}] 名前未発見:`, notFound);
  console.log(`[回収/${label}] ${day}日(${col}列) ${updates.length}件 書き込み完了:`, kaishuu.filter(k => !notFound.includes(k.name)));
  return kaishuu.filter(k => !notFound.includes(k.name));
}

// ── 列インデックス(0始まり) → 列文字（A,B...Z,AA,AB...） ─────
function idxToCol(idx) {
  if (idx < 26) return String.fromCharCode(65 + idx);
  const first  = Math.floor((idx - 26) / 26);
  const second = (idx - 26) % 26;
  return String.fromCharCode(65 + first) + String.fromCharCode(65 + second);
}

// SBシート: 1日のブロック開始列インデックス = T(19)、1日あたり11列
function sbDayStartIdx(day) {
  return 19 + (day - 1) * 11;
}

// ── 日報・給料シートから SB データ読み取り ───────────────────
async function readSbData(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'給料'!A1:I30",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r.data.values || [];

  // 行1はヘッダー、行2以降でA列に名前がある行のみ抽出
  const data = [];
  for (const row of rows.slice(1)) {
    const name = row[0];
    if (!name || typeof name !== 'string' || name.trim() === '') continue;
    // B〜I (index 1〜8) の8値
    const vals = [1,2,3,4,5,6,7,8].map(i => {
      const v = Number(row[i]);
      return isNaN(v) ? 0 : v;
    });
    data.push({ name: name.trim(), vals });
  }
  return data;
}

// ── 月報・SBシートへ書き込み ──────────────────────────────
async function writeSbData(jstDate, sbData, sheetId, label) {
  if (sbData.length === 0) {
    console.log(`[SB/${label}] 書き込み対象なし`);
    return [];
  }

  const day = jstDate.getUTCDate();
  const startIdx = sbDayStartIdx(day);
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // SBシートのA列（名前リスト）を取得
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: "'SB'!A1:A200",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const aCol = (res.data.values || []).map(r => normalizeName((r[0] || '').toString()));

  const updates = [];
  const notFound = [];

  for (const item of sbData) {
    const rowIdx = aCol.findIndex((name, i) => i > 0 && name === item.name);
    if (rowIdx === -1) {
      notFound.push(item.name);
      continue;
    }
    const row = rowIdx + 1; // 1-indexed
    // B〜H の7値を対応列（startIdx〜startIdx+6）に書き込む
    item.vals.forEach((v, offset) => {
      const col = idxToCol(startIdx + offset);
      updates.push({ range: `'SB'!${col}${row}`, values: [[v]] });
    });
  }

  if (updates.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: { valueInputOption: 'RAW', data: updates },
    });
  }

  if (notFound.length > 0) console.log(`[SB/${label}] 名前未発見:`, notFound);
  const written = sbData.filter(k => !notFound.includes(k.name));
  console.log(`[SB/${label}] ${day}日(${idxToCol(startIdx)}列〜) ${written.length}件 書き込み完了`);
  return written;
}

// ── 日報・給料シートからふわもこSBデータ読み取り（K=名前, L〜S=データ）
async function readFuwamokoSbData(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'給料'!K1:S30",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r.data.values || [];

  // 行1はヘッダー、行2以降でK列に名前がある行のみ抽出
  const data = [];
  for (const row of rows.slice(1)) {
    const name = row[0];
    if (!name || typeof name !== 'string' || name.trim() === '') continue;
    if (name.includes('合計')) continue; // 集計行（給与合計等）はスキップ
    // L〜S (index 1〜8) の8値
    const vals = [1,2,3,4,5,6,7,8].map(i => {
      const v = Number(row[i]);
      return isNaN(v) ? 0 : v;
    });
    data.push({ name: name.trim(), vals });
  }
  return data;
}

// ── 日報・手持ちシートからN19:O25（雑費）読み取り ─────────────────
// N列に項目名があり、O列に金額がある行を取得
// 戻り値: [{name: "ガソリン", amount: -4000}, ...] （O×1000、符号反転済み）
async function readZatsuhi(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'手持ち'!N19:O25",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r.data.values || [];

  const result = [];
  for (const row of rows) {
    const name   = (row[0] || '').toString().trim();
    const oVal   = Number(row[1]);
    if (!name || isNaN(oVal) || oVal === 0) continue;
    const amount = oVal * 1000; // ×1000（符号そのまま）
    result.push({ name, amount });
  }
  return result;
}

// ── 月報・経費シートへ雑費書き込み ────────────────────────────
// 交通費と同様に C=日、D=名前、H=金額 の形式で空き行に追記
async function writeZatsuhi(jstDate, zatsuhiItems, sheetId, label) {
  if (!zatsuhiItems || zatsuhiItems.length === 0) {
    console.log(`[雑費/${label}] 書き込み対象なし`);
    return [];
  }

  const day = jstDate.getUTCDate();
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // 経費シートのC・D・H列を読んで空き行と重複チェック
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: "'経費'!C1:H300",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const existingRows = res.data.values || [];

  const existingSet = new Set();
  let startRow = -1;
  for (let i = 3; i < existingRows.length; i++) {
    const row = existingRows[i];
    const c = row?.[0], d = row?.[1], h = row?.[5];
    if (c === '' || c === undefined || c === null) {
      if (startRow === -1) startRow = i + 1;
    } else {
      existingSet.add(`${c}|${d}|${h}`);
    }
  }
  if (startRow === -1) startRow = existingRows.length + 1;

  // 重複を除外
  const newItems = zatsuhiItems.filter(item => {
    const key = `${day}|${item.name}|${item.amount}`;
    if (existingSet.has(key)) {
      console.log(`[雑費/${label}] スキップ(重複): ${day}日 ${item.name} ${item.amount}円`);
      return false;
    }
    return true;
  });

  if (newItems.length === 0) {
    console.log(`[雑費/${label}] 書き込み対象なし（全件重複）`);
    return [];
  }

  const updates = newItems.map((item, idx) => {
    const row = startRow + idx;
    return [
      { range: `'経費'!C${row}`, values: [[day]]        },
      { range: `'経費'!D${row}`, values: [[item.name]]  },
      { range: `'経費'!H${row}`, values: [[item.amount]] },
    ];
  }).flat();

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: { valueInputOption: 'RAW', data: updates },
  });

  console.log(`[雑費/${label}] ${day}日 ${newItems.length}件 書き込み完了(${startRow}行目〜):`, newItems);
  return newItems;
}

// ── 日報・手持ちシートからバンスデータ読み取り ──────────────────
// T列（バンス）に値がある行のK列名前を取得し、"{name}バンス" として返す
// 戻り値: [{name: "山本バンス", amount: -2000}, ...]
async function readBansu(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'手持ち'!K42:T80",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r.data.values || [];

  const result = [];
  for (const row of rows.slice(1)) { // 1行目はヘッダーなのでスキップ
    const name  = row[0];  // K列 = 名前
    const bansu = row[9];  // T列 = バンス (K=0, L=1, ..., T=9)
    if (!name || typeof name !== 'string' || name.trim() === '') continue;
    const amount = Number(bansu);
    if (isNaN(amount) || amount === 0) continue;
    result.push({ name: name.trim() + 'バンス', amount });
  }
  return result;
}

// ── 月報・回収シートへバンス書き込み ──────────────────────────
// D列で "{name}バンス" を検索し、日付列にamountをそのまま（符号反転なし）書き込む
async function writeBansu(jstDate, bansuItems, sheetId, label) {
  if (!bansuItems || bansuItems.length === 0) {
    console.log(`[バンス/${label}] 書き込み対象なし`);
    return [];
  }

  const day = jstDate.getUTCDate();
  const col = dayToKaishuuCol(day);
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // D列の名前リストを取得
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: "'回収'!D1:D200",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const dCol = (res.data.values || []).map(r => (r[0] || '').toString().trim());

  const updates = [];
  const notFound = [];

  for (const item of bansuItems) {
    const rowIdx = dCol.findIndex((name, i) => i > 0 && name === item.name);
    if (rowIdx === -1) {
      notFound.push(item.name);
      continue;
    }
    const row = rowIdx + 1; // 1-indexed
    updates.push({ range: `'回収'!${col}${row}`, values: [[-item.amount]] }); // 符号反転（+→-、-→+）
  }

  if (updates.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: { valueInputOption: 'RAW', data: updates },
    });
  }

  if (notFound.length > 0) console.log(`[バンス/${label}] 名前未発見:`, notFound);
  console.log(`[バンス/${label}] ${day}日(${col}列) ${updates.length}件 書き込み完了:`, bansuItems.filter(k => !notFound.includes(k.name)));
  return bansuItems.filter(k => !notFound.includes(k.name));
}

// ── 日報・売上シートからカード/PayPay読み取り ─────────────────
// 戻り値: { crea: {card, paypay}, fuwamoko: {card, paypay} }
async function readPaymentData(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'売上'!AR2:AZ64",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r.data.values || [];

  const calc = (startRow, endRow) => {
    let card = 0, paypay = 0;
    for (let i = startRow - 2; i <= endRow - 2; i++) { // -2 because AR2=index0
      const row = rows[i];
      if (!row) continue;
      const ar = (row[0] || '').toString().toLowerCase();
      const az = Number(row[8]);
      if (isNaN(az) || az === 0) continue;
      if (ar.includes('paypay') || ar.includes('ペイペイ')) paypay += az * 1000;
      else if (ar.includes('カード')) card += az * 1000;
    }
    return { card, paypay };
  };

  return {
    crea:     calc(3, 46),   // CREA: 行3〜46
    fuwamoko: calc(48, 64),  // ふわもこ: 行48〜64
  };
}

// ── 月報・カード/PayPay書き込み ────────────────────────────
async function writePaymentData(jstDate, payment, sheetId, label) {
  const { card, paypay } = payment;
  if (card === 0 && paypay === 0) {
    console.log(`[支払/${label}] カード・PayPayなし`);
    return;
  }

  const col = dayToCol(jstDate.getUTCDate());
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const updates = [];
  if (card   > 0) updates.push({ range: `'売上'!${col}24`, values: [[card]]   }); // カード
  if (paypay > 0) updates.push({ range: `'売上'!${col}25`, values: [[paypay]] }); // PayPay

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: { valueInputOption: 'RAW', data: updates },
  });

  console.log(`[支払/${label}] ${jstDate.getUTCDate()}日(${col}列) カード:${card} PayPay:${paypay}`);
}

// ── 経費シートのH列合計を売上シート23行に書き込み ──────────
async function writeKeihiTotal(jstDate, sheetId, label) {
  const day = jstDate.getUTCDate();
  const col = dayToCol(day);
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // 経費シートのC列（日）とH列（金額）を取得
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: "'経費'!C1:H300",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = res.data.values || [];

  let total = 0;
  for (const row of rows) {
    if (Number(row[0]) === day) {
      const h = Number(row[5]); // H列はC列から5つ右（index5）
      if (!isNaN(h)) total += h;
    }
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `'売上'!${col}23`,
    valueInputOption: 'RAW',
    requestBody: { values: [[total]] },
  });

  console.log(`[経費合計/${label}] ${day}日(${col}23) 合計:${total}`);
  return total;
}

// ── 日報・手持ちシートから人件費データ読み取り ──────────────────
// 手持ちシートのO35:O41に値がある行のL列名前を取得
// 戻り値: [{name: "すい", hours: 11}, ...]
async function readJinkenhi(fileId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: fileId,
    range: "'手持ち'!L35:O41",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = r.data.values || [];

  const result = [];
  for (const row of rows) {
    const name  = (row[0] || '').toString().trim();
    const hours = Number(row[3]);
    if (!name || isNaN(hours) || hours === 0) continue;
    result.push({ name, hours });
  }
  return result;
}

// ── 月報・人件費シートへ書き込み ──────────────────────────────
// 1行目ヘッダーの「○日」列を検索し、D列名前と照合して書き込む
async function writeJinkenhi(jstDate, jinkenItems, sheetId, label) {
  if (!jinkenItems || jinkenItems.length === 0) {
    console.log(`[人件費/${label}] 書き込み対象なし`);
    return [];
  }

  const day = jstDate.getUTCDate();
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // ヘッダー行 + 名前D列を一括取得
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: "'人件費'!A1:AK50",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = res.data.values || [];

  // 1行目から日付列インデックスを取得
  const header = rows[0] || [];
  const colIdx = header.findIndex(h => h === `${day}日`);
  if (colIdx === -1) {
    console.log(`[人件費/${label}] ${day}日 の列が見つかりません`);
    return [];
  }
  const colStr = colIdx < 26
    ? String.fromCharCode(65 + colIdx)
    : 'A' + String.fromCharCode(65 + colIdx - 26);

  // D列(index3)から名前→行番号マップを作成
  const nameMap = {};
  rows.forEach((row, i) => {
    const name = (row[3] || '').toString().trim();
    if (name) nameMap[name] = i + 1;
  });

  const updates = [];
  const notFound = [];
  for (const item of jinkenItems) {
    const row = nameMap[item.name];
    if (!row) { notFound.push(item.name); continue; }
    updates.push({ range: `'人件費'!${colStr}${row}`, values: [[item.hours]] });
  }

  if (updates.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: { valueInputOption: 'RAW', data: updates },
    });
  }

  if (notFound.length > 0) console.log(`[人件費/${label}] 名前未発見:`, notFound);
  console.log(`[人件費/${label}] ${day}日(${colStr}列) ${updates.length}件 書き込み完了`);
  return jinkenItems.filter(k => !notFound.includes(k.name));
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

  // 当月の月報シートIDをフォルダから動的に取得（見つからなければフォールバック）
  const creaGeppoId  = (await findGeppoSheetId(y, m, 'C')) || GEPPO_SHEET_ID;
  const fuwaGeppoId  = (await findGeppoSheetId(y, m, 'F')) || FUWAMOKO_GEPPO_ID;

  const fileId = await findNippoFileId(jstDate);
  if (!fileId) throw new Error(`日報ファイルが見つかりません: ${label}`);

  const data        = await readNippoData(fileId);
  const fuwaData    = await readFuwamokoData(fileId);
  const keihi       = await readKeihi(fileId);
  const kaishuu     = await readKaishuu(fileId);
  const bansuItems  = await readBansu(fileId);
  const zatsuhiItems = await readZatsuhi(fileId);
  // 新規名前を月報SB・回収シートに自動登録（書き込み前に実行）
  const sbData         = await readSbData(fileId);
  const fuwamokoSbData = await readFuwamokoSbData(fileId);
  await autoRegisterNames(sbData.map(d => d.name),           creaGeppoId, 'SB',   'A', 9, 'CREA-SB');
  await autoRegisterNames(fuwamokoSbData.map(d => d.name),   fuwaGeppoId, 'SB',   'A', 9, 'ふわもこ-SB');
  await autoRegisterNames(kaishuu.crea.map(d => d.name),     creaGeppoId, '回収', 'D', 2, 'CREA-回収');
  await autoRegisterNames(kaishuu.fuwamoko.map(d => d.name), fuwaGeppoId, '回収', 'D', 2, 'ふわもこ-回収');

  const result      = await writeToGeppo(jstDate, data,     creaGeppoId, 'CREA');
  const fuwaResult  = await writeToGeppo(jstDate, fuwaData, fuwaGeppoId, 'ふわもこ');
  const keihiCrea     = await writeKeihi(jstDate, keihi.crea,     creaGeppoId, 'CREA');
  const keihiFuwamoko = await writeKeihi(jstDate, keihi.fuwamoko, fuwaGeppoId, 'ふわもこ');
  const kaishuuCrea     = await writeKaishuu(jstDate, kaishuu.crea,     creaGeppoId, 'CREA');
  const kaishuuFuwamoko = await writeKaishuu(jstDate, kaishuu.fuwamoko, fuwaGeppoId, 'ふわもこ');
  const bansuResult   = await writeBansu(jstDate, bansuItems, creaGeppoId, 'CREA');
  const zatsuhiResult = await writeZatsuhi(jstDate, zatsuhiItems, creaGeppoId, 'CREA');
  const sbCrea     = await writeSbData(jstDate, sbData,         creaGeppoId, 'CREA');
  const sbFuwamoko = await writeSbData(jstDate, fuwamokoSbData, fuwaGeppoId, 'ふわもこ');
  const payment = await readPaymentData(fileId);
  await writePaymentData(jstDate, payment.crea,     creaGeppoId, 'CREA');
  await writePaymentData(jstDate, payment.fuwamoko, fuwaGeppoId, 'ふわもこ');
  await writeKeihiTotal(jstDate, creaGeppoId, 'CREA');
  await writeKeihiTotal(jstDate, fuwaGeppoId, 'ふわもこ');
  const jinkenItems  = await readJinkenhi(fileId);
  const jinkenResult = await writeJinkenhi(jstDate, jinkenItems, creaGeppoId, 'CREA');

  return { label, ...result, fuwamoko: fuwaResult, keihi: { crea: keihiCrea, fuwamoko: keihiFuwamoko }, kaishuu: { crea: kaishuuCrea, fuwamoko: kaishuuFuwamoko }, bansu: bansuResult, zatsuhi: zatsuhiResult, sb: { crea: sbCrea, fuwamoko: sbFuwamoko }, jinkenhi: jinkenResult };
}

module.exports = { syncNippoToGeppo, dayToCol, findNippoFileId };
