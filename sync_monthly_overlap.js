#!/usr/bin/env node
/**
 * C-040 v6.2: 月次SSの「翌月7日列」⇔ 翌月SS本体「1〜7日」の双方向同期
 *
 * 各月SS末尾の「翌月7日列」(例: 4月SSの 5/1〜5/7) と翌月SS本体の 1〜7日列を
 * セル単位で比較し、変更検知して双方向に同期する。
 *
 * 競合解決: 状態ファイル(.sync_state/monthly_overlap_state.json)に直前同期時の
 * 値を記録し、変わった方を勝者にする。両方変わった場合は dst (翌月SS本体) を勝者。
 * 初回(state無し)は dst を master とみなして src を更新。
 *
 * 設置: /app/chatbot/sync_monthly_overlap.js
 * 状態: /app/chatbot/.sync_state/monthly_overlap_state.json
 * ログ: /app/chatbot/_sync_logs/sync_monthly_overlap.log
 * cron: * * * * * /usr/bin/node /app/chatbot/sync_monthly_overlap.js >> /app/chatbot/_sync_logs/sync_monthly_overlap.log 2>&1
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

const KEY_FILE = '/root/.config/chatbot-service-account.json';
const STATE_FILE = '/app/chatbot/.sync_state/monthly_overlap_state.json';

const SHEETS = {
  '2026-04': '1y33e0FlbS2R9d-iMQqaSm29rtYtl1JIcR1PpERwG2W4',
  '2026-05': '1XTmmkZP6k6PIhZ7RPHD75ClC_gfFA11U960wvCvEuB0',
  '2026-06': '1yizaAM_aQFaepv0kYlKBZY7uDUDsJA_9rwSn7hrDqKQ',
  '2026-07': '1k4FVGkUkqR1HUaroC0bH-REaAy8pEM3uFjO3b202e9o',
  '2026-08': '1en-dlxUDLmSEmSPp2mCBgYDYfDfmfKCf7GlPFAN8GSg',
  '2026-09': '1UeVn4llROiPOnISZ86KN3hXp0f4_RCARL07iEflBDNI',
  '2026-10': '19FbRRmAhQONKhsTK3jeHLJV_WLTcpmiP1n64ykAXrOM',
  '2026-11': '1XalqXt1sfswmPasUY2jMtTXduA6VR7JE-ieBDFKX06I',
  '2026-12': '1tuTfLTt1D1yXp8e86IICYkf4iAo9yHlJpr1OVNfZcQY',
};

const PAIRS = [
  { src: '2026-04', dst: '2026-05' },
  { src: '2026-05', dst: '2026-06' },
  { src: '2026-06', dst: '2026-07' },
  { src: '2026-07', dst: '2026-08' },
  { src: '2026-08', dst: '2026-09' },
  { src: '2026-09', dst: '2026-10' },
  { src: '2026-10', dst: '2026-11' },
  { src: '2026-11', dst: '2026-12' },
];

const STORES = ['CREA', 'ふわもこ'];
const AGGREGATE_MARKERS = [
  '📊', '🏪', '📅',
  '出勤人数', '合計',
  '店欠', '前欠', '当欠', '事前欠勤', '当日欠勤', '店都合',
];

// ============================================================
// utility
// ============================================================
function ts() {
  return new Date().toISOString().replace('T', ' ').replace(/\.\d+Z$/, '');
}
function log(msg) {
  console.log(`[${ts()}] ${msg}`);
}

function normalizeName(raw) {
  if (!raw) return '';
  let n = String(raw).replace(/\s|　/g, '');
  n = n.replace(/[（(][^）)]*[）)]/g, '');
  n = n.replace(/\d+\/\d+[A-Za-z]*$/, '');
  for (let i = 0; i < 3; i++) {
    n = n.replace(/\d+$/, '').replace(/[A-Za-z]+$/, '');
  }
  return n;
}

function isAggregateName(name) {
  return AGGREGATE_MARKERS.some(m => name.includes(m));
}

function loadState() {
  try {
    return JSON.parse(fs.readFileSync(STATE_FILE, 'utf-8'));
  } catch {
    return {};
  }
}
function saveState(state) {
  fs.mkdirSync(path.dirname(STATE_FILE), { recursive: true });
  fs.writeFileSync(STATE_FILE, JSON.stringify(state, null, 2), 'utf-8');
}

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: KEY_FILE,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

// ============================================================
// header parsing: header row → array of {colIdx, day, monthHint}
// ============================================================
function parseHeaderRow(headerRow) {
  // Returns [{ colIdx, day, monthSwitch?: number }] for each cell that has a recognizable day.
  // - Day cells like '1金', '2土', '15水'
  // - Month switch cells like '6月→1月', '6月→1', '5月→1金'
  const out = [];
  for (let i = 0; i < headerRow.length; i++) {
    const cell = String(headerRow[i] || '').trim();
    if (!cell) continue;
    // month switch
    const ms = cell.match(/(\d+)月[→\s]*(\d+)/);
    if (ms) {
      out.push({ colIdx: i, day: parseInt(ms[2], 10), monthSwitch: parseInt(ms[1], 10) });
      continue;
    }
    // plain day
    const dm = cell.match(/^(\d+)/);
    if (dm) {
      out.push({ colIdx: i, day: parseInt(dm[1], 10) });
    }
  }
  return out;
}

// Build src cols (1〜7 of next month) by scanning header for monthSwitch then 6 following days
// 戻り値: [{day, colIdx}] day=1..7
function findSrcCols(headerRow, dstYM) {
  const headers = parseHeaderRow(headerRow);
  const dstMonth = dstYM.month;
  // find month switch cell that points to dstMonth
  let switchIdx = headers.findIndex(h => h.monthSwitch === dstMonth);
  if (switchIdx < 0) {
    // also accept switch month equal to dstMonth (some sheets use 翌月名)
    return [];
  }
  // collect 7 consecutive headers starting from switch
  const result = [];
  let expectedDay = 1;
  for (let i = switchIdx; i < headers.length && result.length < 7; i++) {
    const h = headers[i];
    if (h.day === expectedDay) {
      result.push({ day: h.day, colIdx: h.colIdx });
      expectedDay++;
    }
  }
  return result;
}

// Build dst cols (1〜7 of own month) by scanning header for first 7 day cells
function findDstCols(headerRow) {
  const headers = parseHeaderRow(headerRow);
  const result = [];
  for (let i = 0; i < headers.length && result.length < 7; i++) {
    const h = headers[i];
    if (h.day === result.length + 1 && !h.monthSwitch) {
      result.push({ day: h.day, colIdx: h.colIdx });
    }
  }
  return result;
}

function buildStaffMap(rows) {
  // rows: 2D array. row 0 = title, row 1 = header, row 2+ = staff data
  // returns { normalizedName: rowIdx }
  const map = {};
  for (let r = 2; r < rows.length; r++) {
    const raw = String((rows[r] && rows[r][0]) || '').trim();
    if (!raw) continue;
    if (isAggregateName(raw)) continue;
    const norm = normalizeName(raw);
    if (!norm) continue;
    if (norm in map) continue; // first occurrence wins (avoid duplicates)
    map[norm] = r;
  }
  return map;
}

function parseYM(ymStr) {
  const [y, m] = ymStr.split('-').map(s => parseInt(s, 10));
  return { year: y, month: m };
}

// Convert col idx (0-based) to A1 column letters (A, B, ..., Z, AA, AB...)
function colIdxToA1(idx) {
  let s = '';
  let n = idx;
  while (true) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
    if (n < 0) break;
  }
  return s;
}

async function applyWrites(sheets, ssId, store, writes) {
  // writes: [{row, col, value}]  row/col are 0-based
  if (!writes.length) return;
  const data = writes.map(w => ({
    range: `${store}!${colIdxToA1(w.col)}${w.row + 1}`,
    values: [[w.value]],
  }));
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: ssId,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data,
    },
  });
}

// ============================================================
// main sync logic
// ============================================================
async function syncPairStore(sheets, pair, store, state, dryRun, verbose) {
  const srcId = SHEETS[pair.src];
  const dstId = SHEETS[pair.dst];
  if (!srcId || !dstId) return;

  const [srcRes, dstRes] = await Promise.all([
    sheets.spreadsheets.values.get({ spreadsheetId: srcId, range: store }),
    sheets.spreadsheets.values.get({ spreadsheetId: dstId, range: store }),
  ]);
  const srcRows = srcRes.data.values || [];
  const dstRows = dstRes.data.values || [];
  if (srcRows.length < 3 || dstRows.length < 3) return;

  const dstYM = parseYM(pair.dst);
  const srcCols = findSrcCols(srcRows[1] || [], dstYM);
  const dstCols = findDstCols(dstRows[1] || []);
  if (srcCols.length === 0 || dstCols.length === 0) {
    log(`skip: ${pair.src}->${pair.dst}/${store} (srcCols=${srcCols.length} dstCols=${dstCols.length})`);
    return;
  }

  const srcStaff = buildStaffMap(srcRows);
  const dstStaff = buildStaffMap(dstRows);

  const stateKey = `${pair.src}->${pair.dst}/${store}`;
  state[stateKey] = state[stateKey] || {};
  const subState = state[stateKey];

  const srcWrites = [];
  const dstWrites = [];
  let conflicts = 0;
  let initials = 0;

  for (const [name, srcRow] of Object.entries(srcStaff)) {
    const dstRow = dstStaff[name];
    if (dstRow === undefined) continue;

    for (let d = 1; d <= 7; d++) {
      const sc = srcCols.find(c => c.day === d);
      const dc = dstCols.find(c => c.day === d);
      if (!sc || !dc) continue;

      const srcVal = String((srcRows[srcRow] && srcRows[srcRow][sc.colIdx]) || '');
      const dstVal = String((dstRows[dstRow] && dstRows[dstRow][dc.colIdx]) || '');
      const skey = `${name}|${d}`;
      const lastVal = subState[skey];

      if (srcVal === dstVal) {
        subState[skey] = srcVal;
        continue;
      }

      if (lastVal === undefined) {
        // 初回: dst を master とみなす
        if (verbose) log(`  INIT ${stateKey} ${name} day=${d}: src="${srcVal}" → "${dstVal}" (dst-as-master)`);
        srcWrites.push({ row: srcRow, col: sc.colIdx, value: dstVal });
        subState[skey] = dstVal;
        initials++;
        continue;
      }

      const srcChanged = srcVal !== lastVal;
      const dstChanged = dstVal !== lastVal;

      if (srcChanged && !dstChanged) {
        if (verbose) log(`  SRC→DST ${stateKey} ${name} day=${d}: dst "${dstVal}" → "${srcVal}"`);
        dstWrites.push({ row: dstRow, col: dc.colIdx, value: srcVal });
        subState[skey] = srcVal;
      } else if (!srcChanged && dstChanged) {
        if (verbose) log(`  DST→SRC ${stateKey} ${name} day=${d}: src "${srcVal}" → "${dstVal}"`);
        srcWrites.push({ row: srcRow, col: sc.colIdx, value: dstVal });
        subState[skey] = dstVal;
      } else {
        // 両方変わった → dst を勝ちに
        log(`CONFLICT ${stateKey} ${name} day=${d} src="${srcVal}" dst="${dstVal}" last="${lastVal}" → dst wins`);
        srcWrites.push({ row: srcRow, col: sc.colIdx, value: dstVal });
        subState[skey] = dstVal;
        conflicts++;
      }
    }
  }

  if (srcWrites.length || dstWrites.length || initials || conflicts) {
    log(`${stateKey}: src_writes=${srcWrites.length} dst_writes=${dstWrites.length} initials=${initials} conflicts=${conflicts}`);
  }

  // apply writes
  if (!dryRun) {
    if (srcWrites.length) await applyWrites(sheets, srcId, store, srcWrites);
    if (dstWrites.length) await applyWrites(sheets, dstId, store, dstWrites);
  }
}

async function main() {
  const argv = process.argv.slice(2);
  const dryRun = argv.includes('--dry-run');
  const verbose = argv.includes('--verbose');
  const onlyPair = argv.find(a => a.startsWith('--pair='));
  const onlyStore = argv.find(a => a.startsWith('--store='));

  if (dryRun) log('DRY_RUN mode: writes will be skipped (state still updated)');

  const sheets = await getSheetsClient();
  const state = loadState();

  let activePairs = PAIRS;
  if (onlyPair) {
    const v = onlyPair.split('=')[1];
    activePairs = PAIRS.filter(p => `${p.src}->${p.dst}` === v);
  }
  let activeStores = STORES;
  if (onlyStore) {
    const v = onlyStore.split('=')[1];
    activeStores = STORES.filter(s => s === v);
  }

  for (const pair of activePairs) {
    for (const store of activeStores) {
      try {
        await syncPairStore(sheets, pair, store, state, dryRun, verbose);
      } catch (e) {
        log(`ERROR pair=${pair.src}->${pair.dst}/${store}: ${e.message}`);
      }
    }
  }
  if (!dryRun) saveState(state);
  else log('DRY_RUN: state not saved');
}

main().catch(e => {
  log(`FATAL: ${e.message}`);
  process.exit(1);
});
