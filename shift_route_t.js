// /app/chatbot/shift_route_t.js
// T-XXX 系コマンドのルーティング & 認証 & シフト反映 (複数日対応版)
// 2026-05-19 C-064派生
// 2026-05-27 v6 凪+すい: 「新規登録 T-XXX 名前」コマンド追加 (C-036 v5.7 T-XXX版)
// 2026-05-29 v9 凪+すい: CREA/ふわもこ (T-001-a/T-001-b) を新規登録コマンドに統合
//   - 「新規登録 CREA 名前」「新規登録 ふわもこ 名前」「新規登録 T-001-a 名前」許容
//   - 👥スタッフ管理 書込列を tabNames から派生 (CREA=A列 / ふわもこ=C列 / その他=A列)
//   - 店舗マスター: T-001-a/T-001-b を folder=1Zkc.../pattern=2026年X月 シフト表/有効 に修正済
//   - シフト本体反映は従来通り shift_route.js (v5.4) に委譲 (parseStoreIdFromText で T-001系は null)
//
// 設計原則:
//   既存 shift_route.js / shift_reflect.js は無傷。T-001系 (CREA/ふわもこ) の通常シフト反映は
//   別経路で従来通り shift_route.js (v5.4) が処理する。本ハンドラは新規登録コマンドと、
//   T-001以外 (例: T-1043 Angel Spa) のシフト反映を扱う。
//
// 入力フォーマット (例):
//   T-1043
//   みお
//
//   18日(月)休み
//   19日(火)20-24受
//   20日(水)19-24受
//
// v6追加: スタッフ新規登録コマンド
//   入力例(1行): 新規登録 T-1043 みお
//   入力例(複数行): 新規登録\nT-1043\nみお
//   動作: 当月以降の全シフト表SS (店舗マスターE列パターン+G列フォルダ) の
//         「👥 スタッフ管理」A列に append (重複チェック・status無効でも許可)
//   返信ポリシー: 成功サイレント、店名NG/認証NG/重複/エラー時のみLINE返信
//
// 処理フロー (シフト反映):
//   1. 1行目から T-XXX (T-001以外) を抽出
//   2. 店舗マスターSS で T-XXX 存在確認
//   3. C-028「👥 LINE登録者」(A=userId / D=所属契約IDカンマ区切り) で認証
//   4. 店舗 status=有効 確認
//   5. shift_reflect_t.reflectShifts で SS書込 (複数日 batch)
//   6. 成功/失敗を LINE返信

const {google} = require("googleapis");
const path = require("path");
const storeMaster = require("./lib/store_master");
const shiftReflect = require("./shift_reflect_t");

// v8: 曜日漢字判定 (CREA shift_reflect.weekdayKanji と同実装)
const _WEEKDAYS = ['日','月','火','水','木','金','土'];
function weekdayKanji(year, month, day) {
  return _WEEKDAYS[new Date(year, month - 1, day).getDay()];
}

// v8 (2026-05-28 すいさん指示): CREA shift_reflect.js の determineInitialMonth と同等。
// 曜日指定あり → 今月/翌月で一致する方
// 曜日なし    → day<=7 && day<今日 → 翌月 / それ以外 → 今月
function determineInitialMonth(firstEntry, baseDate) {
  if (!firstEntry) {
    return { year: baseDate.getFullYear(), month: baseDate.getMonth() + 1 };
  }
  const curYear = baseDate.getFullYear();
  const curMonth = baseDate.getMonth() + 1;
  const todayDay = baseDate.getDate();
  const nextYear = curMonth === 12 ? curYear + 1 : curYear;
  const nextMonth = curMonth === 12 ? 1 : curMonth + 1;

  if (firstEntry.weekday) {
    if (weekdayKanji(curYear, curMonth, firstEntry.day) === firstEntry.weekday) {
      return { year: curYear, month: curMonth };
    }
    if (weekdayKanji(nextYear, nextMonth, firstEntry.day) === firstEntry.weekday) {
      return { year: nextYear, month: nextMonth };
    }
    return { year: curYear, month: curMonth };
  }
  if (firstEntry.day <= 7 && firstEntry.day < todayDay) {
    return { year: nextYear, month: nextMonth };
  }
  return { year: curYear, month: curMonth };
}

const LINE_REGISTER_SS_ID = "19LgvtnN12QGQzqwQgOVpmFckg111C2RzbyjxI2_hkvA";
const LINE_REGISTER_RANGE_PRIMARY = "👥 LINE登録者!A:D";
const LINE_REGISTER_RANGE_FALLBACK = "📱LINE登録!A:D";

const SA_KEY = path.join(process.env.HOME, ".config/chatbot-service-account.json");

let _auth = null;
function _getAuth() {
  if (_auth) return _auth;
  _auth = new google.auth.GoogleAuth({
    keyFile: SA_KEY,
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"]
  });
  return _auth;
}

// 新規登録/スタッフ管理シート書込 用 (RW権限)
let _authRW = null;
function _getAuthRW() {
  if (_authRW) return _authRW;
  _authRW = new google.auth.GoogleAuth({
    keyFile: SA_KEY,
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive"
    ]
  });
  return _authRW;
}

function parseStoreIdFromText(text) {
  if (!text) return null;
  const firstLine = String(text).split(/\r?\n/)[0].trim();
  // [Angel K-022 2026-07-12] トリガー表記ゆれを吸収: #1043 / T#1043 / #T1043 / T-1043 / T1043 を全て受理。
  //   旧 /^#?T-?(\d{3,4})\b/i はT必須で「#1043 出勤」を弾き無視していた(夕方テスト失敗の根本原因)。
  const m = firstLine.match(/^(?:T#-?|#?T-?|#)(\d{3,4})\b/i);
  if (!m) return null;
  const num = m[1].padStart(3, "0");
  if (num === "001") return null;
  return `T-${num}`;
}

// v6/v9: 「新規登録 <T-XXX|店舗名> <名前>」コマンド検出 (line_handler.js isBotCommand から呼ばれる)
// v9 (2026-05-29): 店舗名 (CREA/ふわもこ等) も許容。実際の店舗存在確認は handleStaffRegisterT で行う
function isStaffRegisterCommand(text) {
  if (!text) return false;
  const parts = String(text).trim().split(/[\s　\r\n]+/);
  if (parts[0] !== "新規登録") return false;
  return parts.length >= 3 && !!parts[1] && !!parts[2];
}

// v9: tabNames カラム値から 👥 スタッフ管理 シートの書込列を派生
// CREA=A列 / ふわもこ=C列 / その他 (Angel Spa 等 1SS=1店舗) = A列デフォルト
const STAFF_MGMT_COL_MAP = {
  "CREA": "A",
  "ふわもこ": "C",
};
function getStaffMgmtCol(storeConfig) {
  const tab = (storeConfig && storeConfig.tabNames) || "";
  return STAFF_MGMT_COL_MAP[tab] || "A";
}

// v9: parts[1] を解決して storeConfig を返す。T-XXX形式 or 店舗名(tabNames/storeName) 両対応
async function resolveStoreConfig(token) {
  // T-XXX形式 (T-1043, T-001-a 等 sub-suffix 含む)
  const tm = String(token || "").match(/^#?T-([\dA-Za-z-]+)$/i);
  if (tm) {
    let raw = tm[1];
    // 数字のみなら3桁にpad (T-1043 → T-1043, T-1 → T-001)
    if (/^\d+$/.test(raw)) raw = raw.padStart(3, "0");
    const storeId = `T-${raw}`;
    return await storeMaster.getStoreById(storeId);
  }
  // 店舗名 (tabNames or storeName 一致 / status=有効 のみ)
  const all = await storeMaster.loadAll();
  return all.find(s =>
    s.status === "有効" && (
      s.tabNames === token ||
      s.storeName === token ||
      (s.storeName || "").startsWith(token + " ") ||
      (s.storeName || "").startsWith(token + "(")
    )
  );
}

function normalizeStaffName(s) {
  return String(s || "")
    .replace(/[　\s]/g, "")
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .toLowerCase();
}

// v9.1 (2026-05-29): C-028 認証データのファイル永続化キャッシュ (5分TTL + Quotaエラー時 stale fallback)
// pm2 restart 跨いで保持・Dashboard監視等の他機能がQuota食い続けても認証通る保険
const fs = require("fs");
const AUTH_DISK_CACHE_FILE = "/tmp/c028_auth_cache.json";
function _loadAuthCacheFromDisk() {
  try {
    const j = JSON.parse(fs.readFileSync(AUTH_DISK_CACHE_FILE, "utf-8"));
    if (j && Array.isArray(j.rows)) return {rows: j.rows, at: j.at || 0};
  } catch (_) {}
  return null;
}
function _saveAuthCacheToDisk(rows, at) {
  try { fs.writeFileSync(AUTH_DISK_CACHE_FILE, JSON.stringify({rows, at})); } catch (_) {}
}
const _initialAuthDisk = _loadAuthCacheFromDisk();
let _authRowsCache = _initialAuthDisk ? _initialAuthDisk.rows : null;
let _authRowsCacheAt = _initialAuthDisk ? _initialAuthDisk.at : 0;
const AUTH_CACHE_TTL_MS = 5 * 60 * 1000;

async function _loadLineRegisterRows() {
  const now = Date.now();
  if (_authRowsCache && (now - _authRowsCacheAt) < AUTH_CACHE_TTL_MS) {
    return _authRowsCache;
  }
  const sheets = google.sheets({version: "v4", auth: _getAuth()});
  let rows = null;
  for (const range of [LINE_REGISTER_RANGE_PRIMARY, LINE_REGISTER_RANGE_FALLBACK]) {
    try {
      const resp = await sheets.spreadsheets.values.get({
        spreadsheetId: LINE_REGISTER_SS_ID,
        range
      });
      rows = resp.data.values || [];
      break;
    } catch (_) {}
  }
  if (rows === null) {
    if (_authRowsCache) {
      console.warn(`[shift_route_t] C-028読込失敗、stale cacheでfallback (age=${Math.round((now-_authRowsCacheAt)/1000)}s)`);
      return _authRowsCache;
    }
    return null;
  }
  _authRowsCache = rows;
  _authRowsCacheAt = now;
  _saveAuthCacheToDisk(rows, now);
  return rows;
}

async function checkLineUserAuth(userId, storeId) {
  const rows = await _loadLineRegisterRows();
  if (rows === null) return {allowed: false, reason: "sheet_read_fail"};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const regUserId = (r[0] || "").trim();
    const regName = (r[1] || "").trim();
    const contractIdsRaw = (r[3] || "").trim();
    if (!regUserId || !contractIdsRaw) continue;
    if (regUserId !== userId) continue;
    const contractIds = contractIdsRaw.split(/[,，]/).map(s => s.trim()).filter(Boolean);
    if (contractIds.includes(storeId)) {
      return {allowed: true, regName};
    }
  }
  return {allowed: false, reason: "not_registered"};
}

// v6: 店舗マスターE列パターン (例「Angel Spa 2026年X月 シフト表」) から
//     当月以降のシフトSSを Drive 内で動的に検索する正規表現を構築
function buildSSNameRegex(patternBase) {
  // 1. 正規表現メタ文字エスケープ
  let escaped = patternBase.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  // 2. "X月" → "(\d{1,2})月" (月キャプチャ)
  escaped = escaped.replace("X月", "(\\d{1,2})月");
  // 3. 年も動的化: "2026年" → "(\d{4})年" (将来年対応)
  escaped = escaped.replace(/(\d{4})年/, "(\\d{4})年");
  return new RegExp("^" + escaped + "$");
}

async function listFutureStoreShiftSS(storeConfig) {
  const client = await _getAuthRW().getClient();
  const drive = google.drive({version: "v3", auth: client});
  const folderId = storeConfig.driveFolderId;
  const r = await drive.files.list({
    q: `'${folderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: "files(id,name)",
    pageSize: 100
  });
  const regex = buildSSNameRegex(storeConfig.monthlySsPattern);
  const now = new Date();
  const curY = now.getFullYear();
  const curM = now.getMonth() + 1;
  const out = [];
  for (const f of (r.data.files || [])) {
    const m = f.name.match(regex);
    if (!m) continue;
    // m[1] が年(動的化したので) or 月。pattern内に年が4桁あれば年がm[1]、月がm[2]
    // 動的化していない場合は月のみキャプチャ。判定:
    const y = (m.length >= 3) ? parseInt(m[1], 10) : (new Date().getFullYear());
    const mo = (m.length >= 3) ? parseInt(m[2], 10) : parseInt(m[1], 10);
    if (y > curY || (y === curY && mo >= curM)) {
      out.push({id: f.id, name: f.name, year: y, month: mo});
    }
  }
  out.sort((a, b) => (a.year - b.year) || (a.month - b.month));
  return out;
}

async function registerNewStaffT(storeConfig, name) {
  const client = await _getAuthRW().getClient();
  const sheets = google.sheets({version: "v4", auth: client});

  // v9: 書込列を tabNames から派生 (CREA=A, ふわもこ=C, その他=A)
  const col = getStaffMgmtCol(storeConfig);

  const ssList = await listFutureStoreShiftSS(storeConfig);
  if (ssList.length === 0) return {type: "no_ss"};

  const normInput = normalizeStaffName(name);
  const targets = [];
  const dupMonths = [];
  let staffSheetMissing = false;

  for (const ss of ssList) {
    let rows;
    try {
      const v = await sheets.spreadsheets.values.get({
        spreadsheetId: ss.id,
        range: `👥 スタッフ管理!${col}2:${col}500`,
        valueRenderOption: "FORMATTED_VALUE"
      });
      rows = v.data.values || [];
    } catch (e) {
      console.error(`[StaffRegisterT] ${ss.name} スタッフ管理シート読込失敗: ${e.message}`);
      staffSheetMissing = true;
      continue;
    }
    let dup = false;
    let lastFilledRow = 1; // header is row 1
    for (let i = 0; i < rows.length; i++) {
      const cell = ((rows[i] || [])[0] || "").toString().trim();
      if (!cell) continue;
      if (/^(📊|🚨|📅)/.test(cell)) continue;
      if (normalizeStaffName(cell) === normInput) dup = true;
      lastFilledRow = i + 2;
    }
    if (dup) dupMonths.push(`${ss.month}月`);
    targets.push({ss, nextRow: lastFilledRow + 1});
  }

  if (targets.length === 0 && staffSheetMissing) {
    return {type: "no_staff_sheet"};
  }
  if (dupMonths.length > 0) {
    return {type: "duplicate", name, months: dupMonths};
  }

  const writeMonths = [];
  for (const t of targets) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: t.ss.id,
      range: `👥 スタッフ管理!${col}${t.nextRow}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {values: [[name]]}
    });
    writeMonths.push(`${t.ss.month}月`);
  }
  console.log(`[StaffRegisterT] success: ${storeConfig.storeName} [${col}列]「${name}」→ ${writeMonths.join("・")}`);
  return {type: "success", name, months: writeMonths, col};
}

async function handleStaffRegisterT(event, parts, client) {
  const userId = event.source?.userId;
  const target = event.source?.groupId || userId;
  if (!userId) return true;

  if (parts.length < 3) {
    await client.pushMessage({to: target, messages: [{type: "text", text:
      "❓ スタッフ名が必要です。\n例: 新規登録 CREA みお / 新規登録 T-1043 みお"
    }]}).catch(() => {});
    return true;
  }
  const name = parts.slice(2).join(" ").trim();
  if (!name) return true;

  // v9: parts[1] を resolveStoreConfig で T-XXX/店舗名 両対応で解決
  let storeConfig;
  try {
    storeConfig = await resolveStoreConfig(parts[1]);
  } catch (e) {
    console.error(`[StaffRegisterT] 店舗解決エラー: ${e.message}`);
    await client.pushMessage({to: target, messages: [{type: "text", text:
      `⚠️ 店舗マスター読込に失敗しました (${parts[1]})`
    }]}).catch(() => {});
    return true;
  }
  if (!storeConfig) {
    await client.pushMessage({to: target, messages: [{type: "text", text:
      `❌ 店舗「${parts[1]}」が見つかりません。\n例: CREA / ふわもこ / T-1043`
    }]}).catch(() => {});
    return true;
  }
  const storeId = storeConfig.storeId;

  console.log(`[StaffRegisterT] storeId=${storeId} (${storeConfig.storeName}) name="${name}" userId=${userId}`);

  // C-028認証 (v9: T-001-a/T-001-b は T-001 認証で許可)
  const authStoreId = storeId.startsWith("T-001") ? "T-001" : storeId;
  const auth = await checkLineUserAuth(userId, authStoreId);
  if (!auth.allowed) {
    await client.pushMessage({to: target, messages: [{type: "text", text:
      `❌ このLINEは ${storeId} (${storeConfig.storeName}) への操作許可がありません。\n\n` +
      `#whoami と送信して自分のLINE userId を確認し、\nオーナーへ「${storeId} への登録」を依頼してください。`
    }]}).catch(() => {});
    console.log(`[StaffRegisterT] 認証NG userId=${userId} storeId=${storeId} (auth=${authStoreId}) reason=${auth.reason}`);
    return true;
  }
  console.log(`[StaffRegisterT] 認証OK userId=${userId} storeId=${storeId} (auth=${authStoreId}) (${storeConfig.storeName})`);

  // v6: status「無効」でも許可 (すいさん指示 2026-05-27)
  // ベンリー反映を伴わないSS書込のみなので、有効化前の下準備として使える

  try {
    const result = await registerNewStaffT(storeConfig, name);
    if (result.type === "duplicate") {
      await client.pushMessage({to: target, messages: [{type: "text", text:
        `⚠️ ${storeConfig.storeName}「${name}」は既に登録済みです（${result.months.join("・")}）`
      }]}).catch(() => {});
    } else if (result.type === "no_ss") {
      await client.pushMessage({to: target, messages: [{type: "text", text:
        `❌ 対象のシフト表SSが見つかりません（${storeConfig.storeName} 当月以降）`
      }]}).catch(() => {});
    } else if (result.type === "no_staff_sheet") {
      await client.pushMessage({to: target, messages: [{type: "text", text:
        `❌ ${storeConfig.storeName} の「👥 スタッフ管理」シートが見つかりません`
      }]}).catch(() => {});
    }
    // success: silent (CREA/ふわもこ v5.7 と同じ運用)
  } catch (err) {
    console.error("[StaffRegisterT] エラー:", err.message);
    await client.pushMessage({to: target, messages: [{type: "text", text:
      `❌ 登録エラー: ${err.message}`
    }]}).catch(() => {});
  }
  return true;
}

async function handle(event, text, client) {
  if (!text || typeof text !== "string") return false;

  // ── v6: 新規登録 T-XXX 名前 コマンド ──
  if (isStaffRegisterCommand(text)) {
    const parts = text.trim().split(/[\s　\r\n]+/);
    return await handleStaffRegisterT(event, parts, client);
  }

  const storeId = parseStoreIdFromText(text);
  if (!storeId) return false;

  const userId = event.source?.userId;
  const target = event.source?.groupId || userId;
  if (!userId) return true;

  console.log(`[shift_route_t] storeId=${storeId} userId=${userId}`);

  let storeConfig;
  try {
    storeConfig = await storeMaster.getStoreById(storeId);
  } catch (e) {
    console.error(`[shift_route_t] 店舗マスター読込エラー: ${e.message}`);
    await client.pushMessage({to: target, messages: [{type: "text", text: `⚠️ 店舗マスター読込に失敗しました (${storeId})`}]}).catch(() => {});
    return true;
  }
  if (!storeConfig) {
    await client.pushMessage({to: target, messages: [{type: "text", text:
      `❌ 店舗 ${storeId} は登録されていません。\nオーナーに登録依頼してください。`
    }]}).catch(() => {});
    return true;
  }

  const auth = await checkLineUserAuth(userId, storeId);
  if (!auth.allowed) {
    await client.pushMessage({to: target, messages: [{type: "text", text:
      `❌ このLINEは ${storeId} (${storeConfig.storeName}) への操作許可がありません。\n\n` +
      `#whoami と送信して自分のLINE userId を確認し、\nオーナーへ「${storeId} への登録」を依頼してください。`
    }]}).catch(() => {});
    console.log(`[shift_route_t] 認証NG userId=${userId} storeId=${storeId} reason=${auth.reason}`);
    return true;
  }
  console.log(`[shift_route_t] 認証OK userId=${userId} storeId=${storeId} (${storeConfig.storeName})`);

  if (storeConfig.status !== "有効") {
    await client.pushMessage({to: target, messages: [{type: "text", text:
      `🔧 店舗 ${storeId} (${storeConfig.storeName}) は現在準備中です (status: ${storeConfig.status})。\n` +
      `オーナーに進捗を確認してください。`
    }]}).catch(() => {});
    return true;
  }

  const shiftData = shiftReflect.parseShiftMessage(text);
  if (!shiftData || !shiftData.entries || shiftData.entries.length === 0) {
    await client.pushMessage({to: target, messages: [{type: "text", text:
      `❓ ${storeId} (${storeConfig.storeName}) 認証OK\n\n` +
      `シフト書式が不明です。\n例:\n${storeId}\nスタッフ名\n\n18日(月)休み\n19日(火)20-24受`
    }]}).catch(() => {});
    return true;
  }

  // v8 (2026-05-28 すいさん指示): 月跨ぎ判定を CREA shift_reflect.js の
  // determineInitialMonth ロジックに揃える。
  // - 曜日指定あり: 今月 or 翌月で曜日一致する方を採用
  // - 曜日なし: day<=7 かつ day<今日 → 翌月扱い (例: 5/28 投稿 "1日" → 6/1)
  // - それ以外: 今月
  const now = new Date();
  const { year, month } = determineInitialMonth(shiftData.entries[0], now);

  // [Angel K-022 2026-07-12] CREA式(SS先)方式へ組替:
  //   まず月次シフト表SSへ書込 → 成功時のみ ベンリー(rd04) 反映(出勤タグ時)。
  //   SS書込が失敗(未登録スタッフ等)ならベンリーも見送り、原因をLINE返信する。
  //   旧「ベンリー先即時反映」は廃止(SSを正とする一本フロー化・二重書込回避)。
  const _firstLine = String(text).split(/\r?\n/)[0] || "";
  const _isAngelShukkin = (storeId === "T-1043" && /出勤/.test(_firstLine));

  try {
    const result = await shiftReflect.reflectShifts({
      storeConfig, year, month,
      name: shiftData.name,
      entries: shiftData.entries
    });
    if (result.success) {
      console.log(`[shift_route_t] reflectShifts OK storeId=${storeId} staff=${result.name} count=${result.count}`);
      if (storeId === "T-1043") {
        // [Angel] SS書込成功後にベンリー反映(出勤タグ時のみ)。GHA不使用。
        if (_isAngelShukkin) {
          try {
            const shiftLines = String(text).split(/\r?\n/).slice(1); // タグ行の後(名前行はパーサが日付無しで無視)
            require("/app/chatbot/angel_shift_apply/angel_apply_hook").applyToVenrey(shiftData.name, shiftLines);
            console.log(`[Angel] T-1043 SS反映OK → ベンリー反映 起動 staff=${shiftData.name}`);
          } catch (e) {
            console.error("[Angel] ベンリー反映フック失敗(SS反映は成功):", e.message);
          }
        }
      } else {
        // 他T-XXXX店舗は従来通りGHA dispatch。
        try {
          const sanitized = String(result.name || "").replace(/["\$`\n\r]/g, "");
          if (sanitized) {
            const dispatchScript = "/root/venrey_dispatch_clients.sh";
            require("child_process").exec(
              `${dispatchScript} "${sanitized}" "${storeId}"`,
              (err) => { if (err) console.error("[Shift→Venrey] dispatch err:", err.message); }
            );
            console.log(`[Shift→Venrey] clients dispatched: store=${storeId} staff=${sanitized}`);
          }
        } catch (e) {
          console.error("[Shift→Venrey] exec failed:", e.message);
        }
      }
    } else {
      // SS反映失敗: T-1043はベンリーも見送り。原因をLINE返信(未登録スタッフ等をユーザーへ通知)。
      if (storeId === "T-1043") {
        console.log(`[Angel] SS反映失敗のためベンリー反映も中止 staff=${shiftData.name} err=${result.error}`);
      }
      await client.pushMessage({to: target, messages: [{type: "text", text: `❌ シフト反映失敗\n${result.error}`}]}).catch(() => {});
    }
  } catch (e) {
    console.error("[shift_route_t] reflectShifts エラー:", e.message);
    await client.pushMessage({to: target, messages: [{type: "text", text: `❌ シフト反映エラー: ${e.message}`}]}).catch(() => {});
  }
  return true;
}

module.exports = {handle, parseStoreIdFromText, checkLineUserAuth, isStaffRegisterCommand};
