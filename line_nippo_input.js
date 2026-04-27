/**
 * line_nippo_input.js
 * LINE メッセージから日報ダッシュボードシートに自動入力するモジュール
 *
 * フォーマット:
 *   #T001
 *   日付:4/23
 *   店名:CREA
 *   名前:立花
 *   指名:R
 *   コース:80
 *   オプション:衣装
 *   備考欄:魂
 *   媒体:
 *   番号:09046558038
 *   予約者:K
 *   来店時間:1400
 */
require('dotenv').config();
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// ── 定数 ─────────────────���───────────────────────────────────
const STORES_SHEET_ID  = '19LgvtnN12QGQzqwQgOVpmFckg111C2RzbyjxI2_hkvA';
const FOLDER_SHEET     = '📁 店舗フォルダ';
const DASHBOARD_SHEET  = 'ダッシュボード';
const REGISTER_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycby6kPgHPcn8aaUm96LLmgZUT4IwfTAg2B3Ow17uawzrWNTWiQLx2cmUXg8_BTPyVRgB/exec';

// セルマッピング（項目名 → セル）
const CELL_MAP = {
  '店名':     'C4',
  '名前':     'C5',
  '指名':     'C6',
  'コース':   'C7',
  '媒体':     'C8',
  'オプション': 'C9',
  '備考欄':   'C10',
  '番号':     'C11',
  '予約者':   'C12',
  '来店時間': 'C13',
};

// ── Google 認証 ──────────────────────────────────────────────
function createAuthClient() {
  const credsPath = path.join(process.env.HOME, '.config/gdrive-server-credentials.json');
  const keysPath  = path.join(process.env.HOME, '.config/gcp-oauth.keys.json');

  if (fs.existsSync(credsPath) && fs.existsSync(keysPath)) {
    const oauthKeys   = JSON.parse(fs.readFileSync(keysPath));
    const credentials = JSON.parse(fs.readFileSync(credsPath));
    const installed   = oauthKeys.installed || oauthKeys.web;
    const client = new google.auth.OAuth2(installed.client_id, installed.client_secret, 'http://localhost');
    client.setCredentials({ refresh_token: credentials.refresh_token });
    return client;
  }
  const client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID, process.env.GOOGLE_CLIENT_SECRET, 'http://localhost');
  client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  return client;
}

// ── メッセージ解析 ─────────────��─────────────────────────────
// #T001 で始まるメッセージを日報入力として検知
function isNippoInput(text) {
  return /^#T\d{3}/i.test(text.trim());
}

// メッセージを解析して店舗IDと各項目を抽出
function parseNippoInput(text) {
  const lines = text.trim().split('\n');

  // 1行目: #T001
  const idMatch = lines[0].match(/^#(T[\-]?\d{3,})/i);
  if (!idMatch) return null;
  // T001 → T-001 に正規化
  let storeId = idMatch[1].toUpperCase();
  if (!storeId.includes('-')) {
    storeId = storeId.replace(/^T(\d+)/, 'T-$1');
  }

  const fields = {};
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    // 「項目名:値」または「項目名　値」（コロン/全角コロン/スペース/タブ区切り）
    const match = line.match(/^(.+?)[:\uff1a\s\u3000]+(.*)$/) ||
                  line.match(/^(.+?)[\s\u3000]+(.+)$/);
    if (match) {
      const key = match[1].trim();
      let val = match[2].trim();
      // 半角英字のみの値は大文字に変換（r→R, t→T など）
      if (/^[a-zA-Z]+$/.test(val)) {
        val = val.toUpperCase();
      }
      fields[key] = val;
    }
  }

  return { storeId, fields };
}

// ── Google DriveのURLからフォルダIDを抽出 ───────────────────
function extractFolderId(input) {
  if (!input) return null;
  const trimmed = input.trim();
  // URL形式: https://drive.google.com/drive/folders/XXXXX?...
  const urlMatch = trimmed.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (urlMatch) return urlMatch[1];
  // ID直貼りの場合はそのまま返す
  if (/^[a-zA-Z0-9_-]+$/.test(trimmed)) return trimmed;
  return null;
}

// ── 店舗ID + 店名 → 日報フォルダID 取得 ─────────────────────
// 「📁 店舗フォルダ」シートから (契約ID, 店舗名) でフォルダURLを検索
async function getFolderByStoreAndName(storeId, inputStoreName) {
  const auth   = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: STORES_SHEET_ID,
    range: `'${FOLDER_SHEET}'!A3:C200`,
    valueRenderOption: 'FORMATTED_VALUE',
  });

  const rows = res.data.values || [];
  // 同じ契約IDの店舗一覧を収集
  const storeRows = rows.filter(r => r[0] === storeId);
  if (storeRows.length === 0) {
    throw new Error(`店舗ID ${storeId} が店舗フォルダシートに見つかりません`);
  }

  // 店名で絞り込み
  const matched = storeRows.find(r => r[1] === inputStoreName);
  if (!matched) {
    const available = storeRows.map(r => r[1]).join('、');
    throw new Error(`${storeId} に「${inputStoreName}」は登録されていません（登録店舗: ${available}）`);
  }

  const folderId = extractFolderId(matched[2]);
  if (!folderId) {
    throw new Error(`${storeId}/${inputStoreName} の日報フォルダURLが未設定です`);
  }

  return { folderId, storeName: matched[1] };
}

// ── 日付文字列を正規化 ─────────────────���────────────────────
function normalizeDateForSearch(dateStr) {
  if (!dateStr) return null;

  // JST基準で年を取得
  const now = new Date();
  const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const year = jst.getUTCFullYear();

  // "4/23" → "2026年4月23日"
  let m = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return `${year}年${parseInt(m[1])}月${parseInt(m[2])}日`;

  // "4月23日" → "2026年4月23日"
  m = dateStr.match(/^(\d{1,2})月(\d{1,2})日?$/);
  if (m) return `${year}年${parseInt(m[1])}月${parseInt(m[2])}日`;

  // "2026年4月23日" そのまま
  if (/^\d{4}年/.test(dateStr)) return dateStr;

  return dateStr;
}

// ��─ フォルダ内から日付でスプレッドシートを特定 ──────────────
async function findSpreadsheetInFolder(folderId, dateStr) {
  const auth  = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const searchDate = normalizeDateForSearch(dateStr);
  if (!searchDate) throw new Error('日付が指定されてい���せん');

  console.log(`[日報入力] フォルダ ${folderId} で "${searchDate}" を検索`);

  const res = await drive.files.list({
    q: `'${folderId}' in parents and name contains '${searchDate}' and mimeType='application/vnd.google-apps.spreadsheet'`,
    fields: 'files(id, name)',
    orderBy: 'createdTime desc',
    pageSize: 5,
  });

  const files = res.data.files || [];
  console.log(`[日報入力] 検索結果: ${files.length}件`, files.map(f => f.name).join(', '));

  if (!files.length) throw new Error(`${searchDate} のスプレッドシートが見つかりません`);

  // ダッシュボードシートが存在するファイルを優先
  const sheets = google.sheets({ version: 'v4', auth });
  for (const file of files) {
    try {
      const meta = await sheets.spreadsheets.get({ spreadsheetId: file.id });
      const hasDashboard = meta.data.sheets.some(s => s.properties.title === DASHBOARD_SHEET);
      if (hasDashboard) {
        console.log(`[日報入力] ダッシュボードシート確認OK: ${file.name}`);
        return { id: file.id, name: file.name };
      }
      console.log(`[日報入力] ${file.name} にダッシュボードシートなし → スキップ`);
    } catch (e) {
      console.log(`[日報入力] ${file.name} の確認エラー: ${e.message}`);
    }
  }
  throw new Error(`${searchDate} のスプレッドシートに「${DASHBOARD_SHEET}」シートが見つかりません`);
}

// ── ダッシュボードシートに書き込み + C16チェック ─────────────
async function writeToDashboard(spreadsheetId, fields) {
  const auth   = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const data = [];
  for (const [key, cell] of Object.entries(CELL_MAP)) {
    const value = fields[key] ?? '';
    data.push({
      range: `'${DASHBOARD_SHEET}'!${cell}`,
      values: [[value]],
    });
  }

  // データ書き込み
  const res = await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data },
  });
  console.log(`[日報入力] ${res.data.totalUpdatedCells}セル書き込み完了`);

  // C16チェックボックスをON
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${DASHBOARD_SHEET}'!C16`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[true]] },
  });
  console.log('[日報入力] C16チェックボックス → TRUE');

  // GAS WebApp 即時登録エンドポイントへ通知
  try {
    const r = await fetch(REGISTER_WEBAPP_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'register', ssId: spreadsheetId }),
    });
    const txt = await r.text();
    console.log(`[日報入力] GAS即時登録: ${r.status} ${txt}`);
  } catch (e) {
    console.log(`[日報入力] GAS即時登録エラー（処理続行）: ${e.message}`);
  }

  return res.data.totalUpdatedCells;
}

// ── キャンセル処理: ダッシュボードA20:R100から該当行のC/D/Eをクリア + G列を備考で上書き ─
// ダッシュボードの列構成（A20起点）:
// A=# B=店名 C=名前 D=指名 E=コース F=オプション G=備考 H=部屋 I=媒体 J=電話番号 K=予約者
async function cancelEntry(spreadsheetId, fields) {
  const auth   = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const name      = (fields['名前'] || '').trim();
  const phone     = (fields['番号'] || '').replace(/[^0-9]/g, '');
  const bikoInput = fields['備考欄'] || '';

  if (!name || !phone) {
    throw new Error(`キャンセルには名前と番号の両方が必要です（名前: "${name}", 番号: "${phone}"）`);
  }

  // ダッシュボード A20:R100 を取得
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${DASHBOARD_SHEET}'!A20:R100`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const rows = res.data.values || [];

  // 名前(C=2) + 電話番号(J=9) のAND一致で検索（下から走査）
  let targetIdx = -1;
  for (let i = rows.length - 1; i >= 0; i--) {
    const rowName  = (rows[i][2] || '').trim();
    const rowPhone = (rows[i][9] || '').replace(/[^0-9]/g, '');
    if (rowName === name && rowPhone === phone) {
      targetIdx = i;
      break;
    }
  }

  if (targetIdx === -1) {
    throw new Error(`キャンセル対象が見つかりません（名前: ${name}、番号: ${phone}）`);
  }

  const targetRow = targetIdx + 20;

  // C/D/E をクリア + G を備考で上書き
  const batchData = [
    { range: `'${DASHBOARD_SHEET}'!C${targetRow}:E${targetRow}`, values: [['', '', '']] },
    { range: `'${DASHBOARD_SHEET}'!G${targetRow}`, values: [[bikoInput]] },
  ];
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data: batchData },
  });
  console.log(`[キャンセル] ダッシュボード 行${targetRow} C/D/Eクリア・G="${bikoInput}"（${name} / ${phone}）`);

  return targetRow;
}

// ── 変更処理: ダッシュボードA20:R100から検索して上書き ───────
// ダッシュボードの列構成（A20起点）:
// A=# B=店名 C=名前 D=指名 E=コース F=オプション G=備考 H=部屋 I=媒体 J=電話番号 K=予約者
const MODIFY_COL_MAP = {
  '店名':     1,  // B列
  '名前':     2,  // C列
  '指名':     3,  // D列
  'コース':   4,  // E列
  'オプション': 5, // F列
  '備考欄':   6,  // G列
  '媒体':     8,  // I列
  '番号':     9,  // J列
  '予約者':   10, // K列
  '来店時間': 14, // O列（開始）
};

async function modifyEntry(spreadsheetId, fields) {
  const auth   = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const name  = fields['名前'] || '';
  const phone = fields['番号'] || '';

  // ダッシュボード A20:R100 を取得
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${DASHBOARD_SHEET}'!A20:R100`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const rows = res.data.values || [];

  // 電話番号(J=9列目) で行を特定（名前が変わる場合があるため番号のみで検索）
  let targetIdx = -1;
  for (let i = rows.length - 1; i >= 0; i--) {
    const rowPhone = (rows[i][9] || '').replace(/[^0-9]/g, '');
    if (rowPhone === phone.replace(/[^0-9]/g, '')) {
      targetIdx = i;
      break;
    }
  }

  if (targetIdx === -1) {
    throw new Error(`変更対象が見つかりません（名前: ${name}、番号: ${phone}）`);
  }

  const targetRow = targetIdx + 20; // シートの行番号（20行目起点）

  // 変更対象の項目をバッチ更新
  const batchData = [];
  for (const [key, colIdx] of Object.entries(MODIFY_COL_MAP)) {
    if (fields[key] !== undefined) {
      const col = String.fromCharCode(65 + colIdx); // 0=A, 1=B, ...
      batchData.push({
        range: `'${DASHBOARD_SHEET}'!${col}${targetRow}`,
        values: [[fields[key]]],
      });
    }
  }

  if (batchData.length === 0) {
    throw new Error('変更するデータがありません');
  }

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data: batchData },
  });
  console.log(`[変更] ダッシュボード 行${targetRow} を${batchData.length}項目上書き（${name} / ${phone}）`);

  return targetRow;
}

// ── メッセージの末尾コマンド判定 ────────────────────────────
function getCommandType(text) {
  const trimmed = text.trim();
  if (/^\s*キャンセル\s*$/m.test(trimmed)) return 'cancel';
  if (/^\s*変更\s*$/m.test(trimmed)) return 'modify';
  return 'input';
}

function removeCommandLine(text, command) {
  return text.replace(new RegExp(`^\\s*${command}\\s*$`, 'm'), '').trim();
}

// ── メイン処理（LINE イベントから呼ばれる）───────────────────
async function processNippoInput(text) {
  const commandType = getCommandType(text);
  // コマンド行を除去してからパース
  let cleanText = text;
  if (commandType === 'cancel') cleanText = removeCommandLine(text, 'キャンセル');
  if (commandType === 'modify') cleanText = removeCommandLine(text, '変更');

  const parsed = parseNippoInput(cleanText);
  if (!parsed) throw new Error('メッセージの解析に失敗しました');

  const { storeId, fields } = parsed;
  console.log(`[日報入力] 店舗ID: ${storeId}, 項目数: ${Object.keys(fields).length}, コマンド: ${commandType}`);

  // 1. 店舗ID + 店名 → フォルダID（整合性チェック込み）
  const inputStoreName = fields['店名'];
  if (!inputStoreName) throw new Error('店名が入力されていません');
  const { folderId, storeName } = await getFolderByStoreAndName(storeId, inputStoreName);
  console.log(`[日報入力] 店舗: ${storeName}, フォルダ: ${folderId}`);

  // 2. 日付でスプレッドシート特定
  const dateStr = fields['日付'] || null;
  const spreadsheet = await findSpreadsheetInFolder(folderId, dateStr);
  console.log(`[日報入力] スプシ: ${spreadsheet.name} (${spreadsheet.id})`);

  // 3. コマンドで分岐
  if (commandType === 'cancel') {
    const cancelRow = await cancelEntry(spreadsheet.id, fields);
    return {
      storeId, storeName, spreadsheetName: spreadsheet.name,
      cellCount: 1, fields, commandType, cancelRow,
    };
  }

  if (commandType === 'modify') {
    const modifyRow = await modifyEntry(spreadsheet.id, fields);
    return {
      storeId, storeName, spreadsheetName: spreadsheet.name,
      cellCount: 1, fields, commandType, modifyRow,
    };
  }

  // 通常入力: ダッシュボードに書き込み
  const cellCount = await writeToDashboard(spreadsheet.id, fields);

  return {
    storeId, storeName, spreadsheetName: spreadsheet.name,
    cellCount, fields, commandType,
  };
}

module.exports = {
  isNippoInput,
  parseNippoInput,
  processNippoInput,
};
