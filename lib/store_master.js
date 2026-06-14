// /app/chatbot/lib/store_master.js
// 店舗マスターSS から店舗情報を読み込むライブラリ
// 2026-05-19 凪 (C-064派生・T-XXX管理土台)
// 2026-05-29 凪+すい: メモリキャッシュ追加 (5分TTL + Quotaエラー時 stale fallback)
//   - Sheets API Quota 枯渇時に店舗マスター読込が失敗してシフト反映が落ちる事故対策
//   - 5分以内の連続呼出はキャッシュ返答 → 1分間Quotaに優しい
//   - API失敗時は期限切れキャッシュでも返す (障害時のgraceful degradation)
// 既存 T-001 (CREA/ふわもこ) はコードハードコードのまま無傷。
// 本ライブラリは T-001 以外の納品先店舗を SS 起点で扱う。

const {google} = require("googleapis");
const path = require("path");
const fs = require("fs");

const STORE_MASTER_SS_ID = "1gtkuCTkOuWOugIJwF7aJqpJVURFgDfV8k5ct9jGw9qE";
const SHEET_RANGE = "店舗一覧!A2:K";
const SA_KEY = path.join(process.env.HOME, ".config/chatbot-service-account.json");
const DISK_CACHE_FILE = "/tmp/store_master_cache.json";

const COLUMNS = {
  storeId: 0, storeName: 1, benryUser: 2, benryPass: 3,
  monthlySsPattern: 4, tabNames: 5, driveFolderId: 6, lineUserIds: 7,
  status: 8, deliveryDate: 9, remarks: 10
};

let _auth = null;
function _getAuth() {
  if (_auth) return _auth;
  _auth = new google.auth.GoogleAuth({
    keyFile: SA_KEY,
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"]
  });
  return _auth;
}

// ファイル永続化キャッシュ (pm2 restart跨いで保持・Quota枯渇でも動く保険)
function _loadCacheFromDisk() {
  try {
    const j = JSON.parse(fs.readFileSync(DISK_CACHE_FILE, "utf-8"));
    if (j && Array.isArray(j.data)) return {data: j.data, at: j.at || 0};
  } catch (_) {}
  return null;
}
function _saveCacheToDisk(data, at) {
  try { fs.writeFileSync(DISK_CACHE_FILE, JSON.stringify({data, at})); } catch (_) {}
}
const _initialDisk = _loadCacheFromDisk();
let _cache = _initialDisk ? _initialDisk.data : null;
let _cacheAt = _initialDisk ? _initialDisk.at : 0;
const CACHE_TTL_MS = 5 * 60 * 1000;

async function loadAll(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && _cache && (now - _cacheAt) < CACHE_TTL_MS) {
    return _cache;
  }
  try {
    const sheets = google.sheets({version: "v4", auth: _getAuth()});
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: STORE_MASTER_SS_ID,
      range: SHEET_RANGE
    });
    const rows = resp.data.values || [];
    const data = rows.map(row => ({
      storeId: row[COLUMNS.storeId] || "",
      storeName: row[COLUMNS.storeName] || "",
      benryUser: row[COLUMNS.benryUser] || "",
      benryPass: row[COLUMNS.benryPass] || "",
      monthlySsPattern: row[COLUMNS.monthlySsPattern] || "",
      tabNames: (row[COLUMNS.tabNames] || "").split(",").map(s => s.trim()).filter(Boolean),
      driveFolderId: row[COLUMNS.driveFolderId] || "",
      lineUserIds: (row[COLUMNS.lineUserIds] || "").split(",").map(s => s.trim()).filter(s => s.startsWith("U")),
      status: row[COLUMNS.status] || "",
      deliveryDate: row[COLUMNS.deliveryDate] || "",
      remarks: row[COLUMNS.remarks] || ""
    }));
    _cache = data;
    _cacheAt = now;
    _saveCacheToDisk(data, now);
    return data;
  } catch (e) {
    if (_cache) {
      console.warn(`[store_master] API失敗、stale cacheでfallback (age=${Math.round((now-_cacheAt)/1000)}s): ${e.message}`);
      return _cache;
    }
    throw e;
  }
}

async function loadActiveStores() {
  const all = await loadAll();
  return all.filter(s => s.status === "有効");
}

async function getStoreById(storeId) {
  const all = await loadAll();
  return all.find(s => s.storeId === storeId) || null;
}

async function getStoreByLineUserId(userId) {
  if (!userId || !userId.startsWith("U")) return null;
  const all = await loadAll();
  return all.find(s => s.lineUserIds.includes(userId)) || null;
}

module.exports = {
  loadAll, loadActiveStores, getStoreById, getStoreByLineUserId,
  STORE_MASTER_SS_ID, COLUMNS
};
