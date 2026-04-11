const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const TARGET_GID = 1873674341;
const SCHEDULE_SPREADSHEET_ID = '10siqLe6B9A7uvNWgRUdHb462RqxCxkGEGMEKTPhY-S8';
const SCHEDULE_GID = 362802905;
const NIPPO_FOLDER_ID = '1isPYyiUqyWXnS1mtpE1_YWJ9QZBTemdJ';
const MIN_COURSE_MIN = 80; // 最短コース（分）

function colLetterToIndex(col) {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result - 1;
}

const STAFF = {
  CA: colLetterToIndex('CA'),
  CB: colLetterToIndex('CB'),
  CD: colLetterToIndex('CD'),
  CF: colLetterToIndex('CF'),
};
const BOOKING = {
  AP: colLetterToIndex('AP'),
  BD: colLetterToIndex('BD'),
  BE: colLetterToIndex('BF'), // 接客終了はBF列
};

let _authClient = null;
function createAuthClient() {
  if (_authClient) return _authClient;

  let client_id, client_secret, refresh_token;

  const credsPath = path.join(process.env.HOME, '.config/gdrive-server-credentials.json');
  const keysPath  = path.join(process.env.HOME, '.config/gcp-oauth.keys.json');

  if (fs.existsSync(credsPath) && fs.existsSync(keysPath)) {
    // 認証ファイルが存在する場合はファイルを優先（access_tokenは使わない：期限切れになるため）
    const oauthKeys = JSON.parse(fs.readFileSync(keysPath));
    const credentials = JSON.parse(fs.readFileSync(credsPath));
    const installed = oauthKeys.installed || oauthKeys.web;
    client_id = installed.client_id;
    client_secret = installed.client_secret;
    refresh_token = credentials.refresh_token;
  } else {
    // ファイルがない場合のみ環境変数を使用
    client_id = process.env.GOOGLE_CLIENT_ID;
    client_secret = process.env.GOOGLE_CLIENT_SECRET;
    refresh_token = process.env.GOOGLE_REFRESH_TOKEN;
  }

  const oauth2Client = new google.auth.OAuth2(client_id, client_secret, 'http://localhost');
  oauth2Client.setCredentials({ refresh_token }); // access_tokenは設定しない（期限切れ防止）
  _authClient = oauth2Client;
  return oauth2Client;
}

// キャッシュ
const cache = {
  availability: { data: null, expiresAt: 0 },
  upcomingSchedule: { data: null, expiresAt: 0 },
  intervalMap: { data: null, expiresAt: 0 },
  todayFileId: { id: null, date: null }, // 今日の日報ファイルID
};

// 数値の時間（小数含む、24超=翌日）から Date を生成
// baseDate: 営業日の基準日（未指定時は現在日）
function hoursToDate(rawVal, baseDate = null) {
  // 基準日の深夜0時を起点にする
  const d = baseDate ? new Date(baseDate) : new Date();
  if (baseDate) d.setHours(0, 0, 0, 0);

  // 100以上は "HHMM" 形式（例: 1530 → 15:30）
  if (rawVal >= 100) {
    const h = Math.floor(rawVal / 100);
    const m = rawVal % 100;
    d.setHours(h, m, 0, 0);
    return d;
  }

  const hours = Math.floor(rawVal);
  const minutes = Math.round((rawVal - hours) * 60);

  if (hours >= 24) {
    d.setDate(d.getDate() + 1);
    d.setHours(hours - 24, minutes, 0, 0);
  } else {
    d.setHours(hours, minutes, 0, 0);
  }
  return d;
}

// 汎用 parseTime: 文字列"HH:MM" / 数値 / 小数 に対応
function parseTime(val, baseDate = null) {
  if (val === null || val === undefined || val === '') return null;

  if (typeof val === 'string') {
    // "HH:MM" 形式
    const match = val.trim().match(/^(\d{1,2}):(\d{2})/);
    if (match) {
      const d = baseDate ? new Date(baseDate) : new Date();
      if (baseDate) d.setHours(0, 0, 0, 0);
      d.setHours(parseInt(match[1], 10), parseInt(match[2], 10), 0, 0);
      return d;
    }
    // 数字のみ文字列（例: "19"）
    const numMatch = val.trim().match(/^(\d+(?:\.\d+)?)/);
    if (numMatch) return hoursToDate(parseFloat(numMatch[1]), baseDate);
    return null;
  }

  if (typeof val === 'number') {
    // Google Sheets小数形式（0〜1）
    if (val > 0 && val < 1) {
      const totalMin = Math.round(val * 24 * 60);
      const d = baseDate ? new Date(baseDate) : new Date();
      if (baseDate) d.setHours(0, 0, 0, 0);
      d.setHours(Math.floor(totalMin / 60), totalMin % 60, 0, 0);
      return d;
    }
    return hoursToDate(val, baseDate);
  }

  return null;
}

// CB列をパース: 時刻と「上」フラグを返す
function parseCB(val, baseDate = null) {
  if (val === null || val === undefined || val === '') return { time: null, hasUpper: false };

  const str = String(val).trim();
  const hasUpper = str.includes('上');

  const match = str.match(/^(\d+(?:\.\d+)?)/);
  if (!match) return { time: null, hasUpper };

  return { time: hoursToDate(parseFloat(match[1]), baseDate), hasUpper };
}

// インターバル（分）をパース
function parseInterval(str) {
  if (!str) return 0;
  const num = parseInt(String(str).replace(/[^0-9]/g, ''), 10);
  return isNaN(num) ? 0 : num;
}

// Date を "HH:MM" にフォーマット（翌日なら 25:00, 26:00... 形式）
function formatTime(date) {
  if (!date) return null;
  const now = new Date();
  const isNextDay = date.getDate() !== now.getDate();
  if (isNextDay) {
    const h = date.getHours() + 24;
    const m = String(date.getMinutes()).padStart(2, '0');
    return `${h}:${m}`;
  }
  return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
}

// 今日の日報ファイルIDをDriveから検索（日付が変わったら再取得）
async function getTodaySpreadsheetId() {
  const biz = getBusinessDate();
  const dateStr = dateToNippoName(biz);
  if (cache.todayFileId.id && cache.todayFileId.date === dateStr) {
    return cache.todayFileId.id;
  }
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const res = await drive.files.list({
    q: `'${NIPPO_FOLDER_ID}' in parents and name contains '${dateStr}'`,
    fields: 'files(id, name)',
    pageSize: 5,
  });
  const file = res.data.files && res.data.files[0];
  if (!file) throw new Error(`今日の日報ファイルが見つかりません: ${dateStr}`);
  cache.todayFileId = { id: file.id, date: dateStr };
  return file.id;
}

async function getAvailability() {
  if (cache.availability.data && Date.now() < cache.availability.expiresAt) {
    return cache.availability.data;
  }
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const spreadsheetId = await getTodaySpreadsheetId();

  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = meta.data.sheets.find(s => s.properties.sheetId === TARGET_GID);
  if (!sheet) throw new Error(`シートが見つかりません (gid=${TARGET_GID})`);
  const sheetName = sheet.properties.title;

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${sheetName}'!AP:CF`,
    valueRenderOption: 'UNFORMATTED_VALUE',
  });

  const rows = res.data.values || [];
  const now = new Date();
  // 営業日の基準日（深夜0〜2時は前日扱い）の深夜0時を起点として時刻を解析する
  const bizDay = getBusinessDate();
  const bizMidnight = new Date(bizDay);
  bizMidnight.setHours(0, 0, 0, 0);
  const baseAP = BOOKING.AP;

  // --- 1. スタッフ情報を収集 ---
  const staffMap = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const nameVal = row[STAFF.CD - baseAP];
    if (!nameVal || typeof nameVal !== 'string') continue;
    const name = nameVal.trim();
    if (name === '' || name === '名前' || name === '源氏名') continue;

    const workStart = parseTime(row[STAFF.CA - baseAP], bizMidnight);
    const cbRaw = row[STAFF.CB - baseAP];
    const { time: workEnd, hasUpper } = parseCB(cbRaw, bizMidnight);
    const interval = parseInterval(row[STAFF.CF - baseAP]);

    if (workStart || workEnd) {
      staffMap[name] = { workStart, workEnd, hasUpper, interval };
    }
  }

  // --- 2. 接客行を収集（AP列に名前がある行）---
  const activeBookings = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const nameVal = row[BOOKING.AP - baseAP];
    if (!nameVal || typeof nameVal !== 'string') continue;
    const name = nameVal.trim();
    if (name === '' || name === '源氏名') continue;

    const serviceStart = parseTime(row[BOOKING.BD - baseAP], bizMidnight);
    const serviceEnd = parseTime(row[BOOKING.BE - baseAP], bizMidnight);
    if (!serviceStart || !serviceEnd) continue;

    const interval = staffMap[name] ? staffMap[name].interval : 0;
    const intervalEnd = new Date(serviceEnd.getTime() + interval * 60 * 1000);

    if (now >= serviceStart && now <= intervalEnd) {
      if (!activeBookings[name] || serviceEnd > activeBookings[name].serviceEnd) {
        activeBookings[name] = { serviceStart, serviceEnd, intervalEnd };
      }
    }
  }

  // --- 3. スタッフごとに空き状況を判定 ---
  const results = [];

  for (const [name, staff] of Object.entries(staffMap)) {
    const { workStart, workEnd, hasUpper, interval } = staff;

    if (!workStart || !workEnd) {
      results.push({ name, status: '本日出勤なし', nextAvailable: null });
      continue;
    }
    if (now < workStart) {
      results.push({ name, status: '出勤前', nextAvailable: formatTime(workStart), workEnd: formatTime(workEnd), hasUpper, start: formatTime(workStart), end: formatTime(workEnd) });
      continue;
    }
    if (now > workEnd) {
      results.push({ name, status: '退勤済み', nextAvailable: null });
      continue;
    }

    // 接客中・インターバル中チェック
    const booking = activeBookings[name];
    if (booking) {
      if (now >= booking.serviceStart && now <= booking.serviceEnd) {
        results.push({
          name,
          status: '接客中',
          serviceEnd: formatTime(booking.serviceEnd),
          nextAvailable: formatTime(booking.intervalEnd),
          start: formatTime(workStart), end: formatTime(workEnd), hasUpper,
        });
      } else {
        results.push({
          name,
          status: 'インターバル中',
          nextAvailable: formatTime(booking.intervalEnd),
          start: formatTime(workStart), end: formatTime(workEnd), hasUpper,
        });
      }
      continue;
    }

    // 空き：受付可能かチェック（最短80分コース）
    if (hasUpper) {
      // (上)あり: 今から80分後がworkEnd以内でないと新規不可
      const minEndTime = new Date(now.getTime() + MIN_COURSE_MIN * 60 * 1000);
      if (minEndTime > workEnd) {
        results.push({
          name,
          status: '受付終了',
          note: `本日の受付は終了しました（${formatTime(workEnd)}終了）`,
          start: formatTime(workStart), end: formatTime(workEnd), hasUpper,
        });
        continue;
      }
      // 受付可能: 最終受付時間 = workEnd - 80分
      const lastAccept = new Date(workEnd.getTime() - MIN_COURSE_MIN * 60 * 1000);
      results.push({
        name,
        status: '空き',
        workEnd: formatTime(workEnd),
        lastAccept: formatTime(lastAccept),
        hasUpper: true,
        start: formatTime(workStart), end: formatTime(workEnd),
      });
    } else {
      // (上)なし: workEndまで来店可能
      results.push({
        name,
        status: '空き',
        workEnd: formatTime(workEnd),
        hasUpper: false,
        start: formatTime(workStart), end: formatTime(workEnd),
      });
    }
  }

  const result = { results, now: formatTime(now) };
  cache.availability.data = result;
  cache.availability.expiresAt = Date.now() + 60 * 1000; // 60秒
  return result;
}

// スケジュール文字列をパース: "11-18上", "1930-25上", "19-24*130" など
function parseScheduleCell(val) {
  if (!val || String(val).trim() === '') return null;
  const str = String(val).trim();
  const hasUpper = str.includes('上');

  // "HHMM-HHMM" or "HH-HH" 形式を抽出（*以降や上は無視）
  const match = str.match(/^(\d{2,4})-(\d{2,4})/);
  if (!match) return null;

  function parseHHMM(s) {
    if (s.length <= 2) return { h: parseInt(s, 10), m: 0 };
    // 3〜4桁: 下2桁が分、残りが時
    const m = parseInt(s.slice(-2), 10);
    const h = parseInt(s.slice(0, -2), 10);
    return { h, m };
  }

  const start = parseHHMM(match[1]);
  const end = parseHHMM(match[2]);

  const startStr = `${String(start.h).padStart(2,'0')}:${String(start.m).padStart(2,'0')}`;
  const endStr = `${String(end.h).padStart(2,'0')}:${String(end.m).padStart(2,'0')}`;

  return { startStr, endStr, hasUpper };
}

// 翌日以降のスケジュールを取得（今日含む翌7日分）
async function getUpcomingSchedule() {
  if (cache.upcomingSchedule.data && Date.now() < cache.upcomingSchedule.expiresAt) {
    return cache.upcomingSchedule.data;
  }
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const meta = await sheets.spreadsheets.get({ spreadsheetId: SCHEDULE_SPREADSHEET_ID });
  const sheet = meta.data.sheets.find(s => s.properties.sheetId === SCHEDULE_GID);
  if (!sheet) throw new Error('スケジュールシートが見つかりません');
  const sheetName = sheet.properties.title;

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SCHEDULE_SPREADSHEET_ID,
    range: `'${sheetName}'!A1:AJ200`,
    valueRenderOption: 'FORMATTED_VALUE',
  });

  const rows = res.data.values || [];
  if (rows.length < 2) return {};

  const dayRow = rows[0]; // 日付番号の行

  // 今日から7日分の列インデックスを特定
  const today = new Date();
  const result = {};

  for (let offset = 1; offset <= 7; offset++) {
    const target = new Date(today);
    target.setDate(today.getDate() + offset);
    const month = target.getMonth() + 1;
    const day = target.getDate();

    // シート名から月を取得（例: "2026年4月" → 4, "12月" → 12）
    const sheetMonthMatch = sheetName.match(/(\d{1,2})月/);
    const sheetMonth = sheetMonthMatch ? parseInt(sheetMonthMatch[1]) : month;

    // dayRow から該当日の列インデックスを探す
    // 月が変わる場合（例: 3月シートで4月2日を探す）は2回目の出現を使う
    let colIdx = -1;
    let count = 0;
    for (let c = 3; c < dayRow.length; c++) {
      if (String(dayRow[c]).trim() === String(day)) {
        count++;
        if (month === sheetMonth && count === 1) { colIdx = c; break; }
        if (month !== sheetMonth && count === 2) { colIdx = c; break; }
        // month !== sheetMonth のときは count===2 を待つ（フォールバックしない）
      }
    }
    if (colIdx === -1) continue; // 該当列なし（次月データがシートにない等）

    const dateKey = `${month}月${day}日`;
    const staffList = [];

    for (let r = 3; r < rows.length; r++) {
      const nameRaw = rows[r][0];
      if (!nameRaw || String(nameRaw).trim() === '') continue;
      // 名前の最初の部分（姓）を抽出
      const name = String(nameRaw).trim().split(/[\s　]/)[0];
      const cellVal = rows[r][colIdx];
      const parsed = parseScheduleCell(cellVal);
      if (parsed) {
        staffList.push({
          name,
          start: parsed.startStr,
          end: parsed.endStr,
          hasUpper: parsed.hasUpper,
        });
      }
    }

    if (staffList.length > 0) result[dateKey] = staffList;
  }

  cache.upcomingSchedule.data = result;
  cache.upcomingSchedule.expiresAt = Date.now() + 5 * 60 * 1000; // 5分
  return result;
}

// 日報から指定日の予約状況を取得
// dateStr: "2026年3月26日" 形式
async function getDailyReportBookings(dateStr) {
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const sheets = google.sheets({ version: 'v4', auth });

  // 日報フォルダからファイルを検索（スペース違い対応）
  const res = await drive.files.list({
    q: `'${NIPPO_FOLDER_ID}' in parents and name contains '${dateStr}'`,
    fields: 'files(id, name)',
    pageSize: 5,
  });
  const file = res.data.files && res.data.files[0];
  if (!file) return { date: dateStr, bookings: [], note: '日報ファイルなし' };

  const meta = await sheets.spreadsheets.get({ spreadsheetId: file.id });
  const sheetName = meta.data.sheets[0].properties.title;

  const data = await sheets.spreadsheets.values.get({
    spreadsheetId: file.id,
    range: `'${sheetName}'!AP:CF`,
    valueRenderOption: 'FORMATTED_VALUE',
  });

  const rows = data.data.values || [];
  const base = 41; // AP列のインデックス
  const bookings = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const name = row[0];           // AP (offset 0)
    const serviceStart = row[55 - base]; // BD (offset 14)
    const serviceEnd = row[57 - base];   // BF (offset 16)
    if (name && String(name).trim() && String(name).trim() !== '源氏名' && serviceStart && String(serviceStart).match(/\d+:\d+/)) {
      bookings.push({
        name: String(name).trim(),
        start: String(serviceStart).trim(),
        end: serviceEnd ? String(serviceEnd).trim() : '',
      });
    }
  }

  return { date: dateStr, bookings };
}

// インターバルスプレッドシートから名前→インターバル分数のマップを取得
const INTERVAL_SPREADSHEET_ID = '1sh6MgPL4k2StMofKuys_5QA0-3JiDbdY98qe8bvAJEo';

async function getIntervalMap() {
  if (cache.intervalMap.data && Date.now() < cache.intervalMap.expiresAt) {
    return cache.intervalMap.data;
  }
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: INTERVAL_SPREADSHEET_ID,
    range: 'P3:S100',
    valueRenderOption: 'FORMATTED_VALUE',
  });

  const rows = res.data.values || [];
  const intervalMap = {};
  const colIntervals = [10, 15, 20, 5]; // P=10分, Q=15分, R=20分, S=5分

  for (const row of rows) {
    for (let i = 0; i < 4; i++) {
      const name = row[i];
      if (name && String(name).trim()) {
        intervalMap[String(name).trim()] = colIntervals[i];
      }
    }
  }

  cache.intervalMap.data = intervalMap;
  cache.intervalMap.expiresAt = Date.now() + 5 * 60 * 1000; // 5分
  return intervalMap;
}

// 日付文字列からDateを生成（例: "2026年3月26日", "明日", "明後日"）
function getBusinessDate(base = new Date()) {
  // VPSはUTCタイムゾーンのためJST(UTC+9)に変換してから判定
  // 営業日: JST午前6時〜翌午前5:59を同日扱い
  const jst = new Date(base.getTime() + 9 * 60 * 60 * 1000);
  if (jst.getUTCHours() < 6) {
    // JST午前6時前 → 前営業日
    const prev = new Date(jst.getTime() - 24 * 60 * 60 * 1000);
    return new Date(prev.getUTCFullYear(), prev.getUTCMonth(), prev.getUTCDate());
  }
  return new Date(jst.getUTCFullYear(), jst.getUTCMonth(), jst.getUTCDate());
}

function parseDateStr(str, baseDate = new Date()) {
  if (!str) return null;
  str = str.trim();
  const biz = getBusinessDate(baseDate);
  if (str === '今日' || str === '本日') return biz;
  if (str === '明日') { const d = new Date(biz); d.setDate(d.getDate() + 1); return d; }
  if (str === '明後日') { const d = new Date(biz); d.setDate(d.getDate() + 2); return d; }
  const m = str.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  const m2 = str.match(/(\d{1,2})月(\d{1,2})日/);
  if (m2) return new Date(baseDate.getFullYear(), parseInt(m2[1]) - 1, parseInt(m2[2]));
  // DD日のみ（当月として解釈）
  const m3 = str.match(/^(\d{1,2})日$/);
  if (m3) return new Date(baseDate.getFullYear(), baseDate.getMonth(), parseInt(m3[1]));
  return null;
}

function dateToNippoName(date) {
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
}

async function warmupCache() {
  console.log('キャッシュ事前取得中...');
  await Promise.all([getAvailability(), getUpcomingSchedule(), getIntervalMap()]);
  console.log('キャッシュ完了');
}

module.exports = { getAvailability, getUpcomingSchedule, getDailyReportBookings, getIntervalMap, parseDateStr, dateToNippoName, warmupCache, getBusinessDate };
