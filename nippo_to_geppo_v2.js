/**
 * nippo_to_geppo_v2.js
 * 新日報構造対応の日報→月報同期
 *
 * - 日報SS: DATA_FOLDER_ID 内の「YYYY年M月D日」で検索
 * - 月報SS: DATA_FOLDER_ID 内の「CREA売上YYYY-M月」「ふわもこ売上YYYY-M月」で検索
 * - 書込先タブ: 売上 / 経費 / 回収 / SB
 */

require('dotenv').config();
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// データフォルダ（日報・月報の両方が入っている）
const DATA_FOLDER_ID = process.env.DATA_FOLDER_ID || '16R1BK5NnvYkH4Eqh6t51OGQl0tXVJ3If';

// 認証
let _authClient = null;
function createAuthClient() {
  if (_authClient) return _authClient;
  const saPath = path.join(process.env.HOME, '.config/chatbot-service-account.json');
  if (fs.existsSync(saPath)) {
    _authClient = new google.auth.GoogleAuth({
      keyFile: saPath,
      scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
      ],
    });
  } else {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
      'http://localhost'
    );
    oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
    _authClient = oauth2Client;
  }
  return _authClient;
}

// ── 列変換 ───────────────────────────────────
function dayToCol(day) {
  const idx = 6 + (day - 1); // G=6 (0-indexed: A=0)
  if (idx < 26) return String.fromCharCode(65 + idx);
  return 'A' + String.fromCharCode(65 + idx - 26);
}

// 人件費タブ用: E=1日（4 + day - 1）
function dayToJinkenCol(day) {
  const idx = 4 + (day - 1); // E=4
  if (idx < 26) return String.fromCharCode(65 + idx);
  return 'A' + String.fromCharCode(65 + idx - 26);
}

function idxToCol(idx) {
  if (idx < 26) return String.fromCharCode(65 + idx);
  const first = Math.floor((idx - 26) / 26);
  const second = (idx - 26) % 26;
  return String.fromCharCode(65 + first) + String.fromCharCode(65 + second);
}

// ── 時刻パーサー ─────────────────────────────
function parseTimeToHour(v) {
  if (v === null || v === undefined || v === '') return NaN;
  if (typeof v === 'number') return v < 2 ? v * 24 : v;
  const s = String(v).trim();
  const hm = s.match(/^(\d{1,2}):(\d{2})$/);
  if (hm) return Number(hm[1]) + Number(hm[2]) / 60;
  const num = s.match(/^(\d+(?:\.\d+)?)/);
  if (num) return Number(num[1]);
  return NaN;
}

function shiftSlotOverlap(u, v) {
  let s = parseTimeToHour(u);
  let e = parseTimeToHour(v);
  if (isNaN(s) || isNaN(e)) return null;
  if (s < 9) s += 24;
  if (e <= s) e += 24;
  const overlaps = (ps, pe) => s < pe && e > ps;
  return {
    am:    overlaps(9,  15),
    pm:    overlaps(15, 21),
    night: overlaps(21, 30),
  };
}

function timeSlot(v) {
  let hh = parseTimeToHour(v);
  if (isNaN(hh)) return null;
  if (hh < 9) hh += 24;
  if (hh >= 9  && hh < 15) return 'am';
  if (hh >= 15 && hh < 21) return 'pm';
  if (hh >= 21 && hh <= 30) return 'night';
  return null;
}

// ── ファイル検索 ─────────────────────────────
async function findNippoByDate(jstDate) {
  const y = jstDate.getUTCFullYear();
  const m = jstDate.getUTCMonth() + 1;
  const d = jstDate.getUTCDate();
  const dateStr = `${y}年${m}月${d}日`;

  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const res = await drive.files.list({
    q: `'${DATA_FOLDER_ID}' in parents and name='${dateStr}' and trashed=false`,
    fields: 'files(id,name)',
    pageSize: 5,
  });
  const files = res.data.files || [];
  const exact = files.find(f => f.name.trim() === dateStr);
  if (!exact) throw new Error(`日報ファイルが見つかりません: ${dateStr}`);
  return { id: exact.id, name: exact.name };
}

async function findGeppoByStoreMonth(storeName, year, month) {
  const name = `${storeName}売上${year}-${month}月`;
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const res = await drive.files.list({
    q: `'${DATA_FOLDER_ID}' in parents and name='${name}' and trashed=false`,
    fields: 'files(id,name)',
    pageSize: 5,
  });
  const files = res.data.files || [];
  if (!files.length) throw new Error(`月報シートが見つかりません: ${name}`);
  return files[0].id;
}

// ── メイン同期関数 ───────────────────────────
async function syncNippoToGeppo(jstDate) {
  if (!jstDate) {
    // JST昨日
    const now = new Date();
    const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    jst.setUTCDate(jst.getUTCDate() - 1);
    jstDate = jst;
  }
  const y = jstDate.getUTCFullYear();
  const m = jstDate.getUTCMonth() + 1;
  const d = jstDate.getUTCDate();
  const label = `${y}年${m}月${d}日`;
  const col = dayToCol(d);

  console.log(`[v2] ${label} 同期開始`);

  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // 日報SS検索
  const nippo = await findNippoByDate(jstDate);
  console.log(`[v2] 日報: ${nippo.name} (${nippo.id})`);

  // 月報SS検索（店舗×月）
  const creaGeppoId = await findGeppoByStoreMonth('CREA', y, m);
  const fuwaGeppoId = await findGeppoByStoreMonth('ふわもこ', y, m);
  console.log(`[v2] CREA月報=${creaGeppoId}, FUWA月報=${fuwaGeppoId}`);

  // ── ダッシュボード取得 ─────────────────────
  const dashRes = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: nippo.id,
    ranges: [
      "'ダッシュボード'!AG3:AG4",
      "'ダッシュボード'!AN3:AN4",
      "'ダッシュボード'!AH7",
      "'ダッシュボード'!AH9",
      "'ダッシュボード'!AF7",
      "'ダッシュボード'!AF9",
      "'ダッシュボード'!U4:V33",
      "'ダッシュボード'!U36:V200",
    ],
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const [agRows, anRows, ah7Row, ah9Row, af7Row, af9Row, creaURows, fuwaURows] =
    dashRes.data.valueRanges.map(v => v.values || []);

  const creaSales = Number(agRows[0]?.[0]) || 0;
  const fuwaSales = Number(agRows[1]?.[0]) || 0;
  const creaHon   = Number(anRows[0]?.[0]) || 0;
  const fuwaHon   = Number(anRows[1]?.[0]) || 0;
  const creaShukkin = Number(ah7Row[0]?.[0]) || 0;
  const fuwaShukkin = Number(ah9Row[0]?.[0]) || 0;
  const creaTaiki   = Number(af7Row[0]?.[0]) || 0;
  const fuwaTaiki   = Number(af9Row[0]?.[0]) || 0;

  const countBySlot = (rows) => {
    let am=0, pm=0, night=0;
    for (const r of rows) {
      const ov = shiftSlotOverlap(r[0], r[1]);
      if (!ov) continue;
      if (ov.am) am++;
      if (ov.pm) pm++;
      if (ov.night) night++;
    }
    return { am, pm, night };
  };
  const creaU = countBySlot(creaURows);
  const fuwaU = countBySlot(fuwaURows);

  // ── 日報_全件 ───────────────────────────
  const zenRes = await sheets.spreadsheets.values.get({
    spreadsheetId: nippo.id,
    range: "'日報_全件'!B2:O1000",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const zenRows = zenRes.data.values || [];
  let creaHonBySlot = { am:0, pm:0, night:0 };
  let fuwaHonBySlot = { am:0, pm:0, night:0 };
  let creaCard = 0, fuwaCard = 0, creaPay = 0, fuwaPay = 0;
  for (const r of zenRows) {
    const store = (r[0] || '').toString().trim();
    const isCrea = store === 'CREA';
    const isFuwa = store.includes('ふわもこ');
    const slot = timeSlot(r[13]);
    if (slot) {
      if (isCrea) creaHonBySlot[slot]++;
      else if (isFuwa) fuwaHonBySlot[slot]++;
    }
    const opt = (r[4] || '').toString();
    const ryokin = Number(r[10]);
    if (!isNaN(ryokin) && ryokin > 0) {
      const yen = ryokin * 1000;
      if (opt.includes('カード')) {
        if (isCrea) creaCard += yen;
        else if (isFuwa) fuwaCard += yen;
      } else if (opt.includes('PayPay') || opt.includes('ペイペイ') || opt.includes('paypay')) {
        if (isCrea) creaPay += yen;
        else if (isFuwa) fuwaPay += yen;
      }
    }
  }

  // ── 売上タブ書込 ────────────────────────
  const writeSales = async (sheetId, label2, vals) => {
    const updates = [
      { range: `'売上'!${col}3`,  values: [[vals.sales]] },
      { range: `'売上'!${col}5`,  values: [[vals.am_hon]] },
      { range: `'売上'!${col}6`,  values: [[vals.pm_hon]] },
      { range: `'売上'!${col}7`,  values: [[vals.night_hon]] },
      { range: `'売上'!${col}8`,  values: [[vals.hon]] },
      { range: `'売上'!${col}9`,  values: [[vals.am_u]] },
      { range: `'売上'!${col}10`, values: [[vals.pm_u]] },
      { range: `'売上'!${col}11`, values: [[vals.night_u]] },
      { range: `'売上'!${col}12`, values: [[vals.shukkin]] },
      { range: `'売上'!${col}14`, values: [[vals.taiki]] },
      { range: `'売上'!${col}24`, values: [[vals.card]] },
      { range: `'売上'!${col}25`, values: [[vals.pay]] },
    ];
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: { valueInputOption: 'RAW', data: updates },
    });
    console.log(`[v2/売上/${label2}] ${col}列 完了`);
  };

  await writeSales(creaGeppoId, 'CREA', {
    sales: creaSales, am_hon: creaHonBySlot.am, pm_hon: creaHonBySlot.pm, night_hon: creaHonBySlot.night,
    hon: creaHon, am_u: creaU.am, pm_u: creaU.pm, night_u: creaU.night, shukkin: creaShukkin,
    taiki: creaTaiki, card: creaCard, pay: creaPay,
  });
  await writeSales(fuwaGeppoId, 'ふわもこ', {
    sales: fuwaSales, am_hon: fuwaHonBySlot.am, pm_hon: fuwaHonBySlot.pm, night_hon: fuwaHonBySlot.night,
    hon: fuwaHon, am_u: fuwaU.am, pm_u: fuwaU.pm, night_u: fuwaU.night, shukkin: fuwaShukkin,
    taiki: fuwaTaiki, card: fuwaCard, pay: fuwaPay,
  });

  // ── CREA 経費シート書込（現金管理 O4:P23 → 経費 C/D/H）
  const genkinRes = await sheets.spreadsheets.values.get({
    spreadsheetId: nippo.id,
    range: "'現金管理'!O4:P23",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const genkinRows = genkinRes.data.values || [];
  const keihi = [];
  for (const row of genkinRows) {
    const name = (row[0] || '').toString().trim();
    const amount = Number(row[1]);
    if (!name || isNaN(amount) || amount === 0) continue;
    keihi.push({ name, amount });
  }

  if (keihi.length > 0) {
    const exRes = await sheets.spreadsheets.values.get({
      spreadsheetId: creaGeppoId,
      range: "'経費'!C1:H300",
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const exRows = exRes.data.values || [];
    const existingSet = new Set();
    let startRow = -1;
    for (let i = 3; i < exRows.length; i++) {
      const row = exRows[i];
      const c = row?.[0], dv = row?.[1], h = row?.[5];
      if (c === '' || c === undefined || c === null) {
        if (startRow === -1) startRow = i + 1;
      } else {
        existingSet.add(`${c}|${dv}|${h}`);
      }
    }
    if (startRow === -1) startRow = exRows.length + 1;

    const newItems = keihi.filter(item => {
      const key = `${d}|${item.name}|${item.amount}`;
      return !existingSet.has(key);
    });

    if (newItems.length > 0) {
      const updates = newItems.map((item, idx) => {
        const row = startRow + idx;
        return [
          { range: `'経費'!C${row}`, values: [[d]]     },
          { range: `'経費'!D${row}`, values: [[item.name]]  },
          { range: `'経費'!H${row}`, values: [[item.amount]] },
        ];
      }).flat();
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: creaGeppoId,
        requestBody: { valueInputOption: 'RAW', data: updates },
      });
      console.log(`[v2/経費/CREA] ${d}日 ${newItems.length}件 書込`);
    }
  }

  // ── 回収シート書込 ────────────────────────
  const kaishuuSrc = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: nippo.id,
    ranges: ["'現金管理'!A5:G34", "'現金管理'!A37:G80"],
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const [creaRows, fuwaRows] = kaishuuSrc.data.valueRanges.map(v => v.values || []);

  const extractKaishuu = (rows) => {
    const items = [];
    for (const r of rows) {
      const name = (r[0] || '').toString().trim();
      if (!name || name.includes('▌')) continue;
      if (name.replace(/\s/g, '').includes('合計')) continue;
      const fn = isNaN(Number(r[5])) ? 0 : Number(r[5]);
      const gn = isNaN(Number(r[6])) ? 0 : Number(r[6]);
      const amount = (fn > 0 && gn >= 0) ? (gn - fn) : (fn + gn);
      if (amount === 0) continue;
      items.push({ name, amount });
    }
    return items;
  };
  const creaKaishuu = extractKaishuu(creaRows);
  const fuwaKaishuu = extractKaishuu(fuwaRows);

  const writeKaishuu = async (sheetId, label2, items) => {
    if (items.length === 0) return;
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: "'回収'!A1:AK200",
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const rows = res.data.values || [];
    const header = rows[0] || [];
    const colIdx = header.findIndex(h => String(h).trim() === `${d}日`);
    if (colIdx === -1) { console.log(`[v2/回収/${label2}] ${d}日 列なし`); return; }
    const colStr = idxToCol(colIdx);

    const MAX_ROW = 46;
    const nameMap = {};
    const emptyRows = [];
    for (let i = 1; i < rows.length; i++) {
      const row = i + 1;
      if (row > MAX_ROW) break;
      const v = rows[i]?.[3];
      const n = (v || '').toString().trim();
      if (n) nameMap[n] = row;
      else emptyRows.push(row);
    }

    const toAdd = items.filter(it => !nameMap[it.name]);
    if (toAdd.length > 0) {
      if (toAdd.length > emptyRows.length) {
        throw new Error(`[v2/回収/${label2}] D2〜D46に空白行不足（必要${toAdd.length}/空${emptyRows.length}）: ${toAdd.map(i=>i.name).join(',')}`);
      }
      const addUpdates = toAdd.map((it, idx) => ({
        range: `'回収'!D${emptyRows[idx]}`,
        values: [[it.name]],
      }));
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: { valueInputOption: 'RAW', data: addUpdates },
      });
      toAdd.forEach((it, idx) => { nameMap[it.name] = emptyRows[idx]; });
      console.log(`[v2/回収/${label2}] 新規登録: ${toAdd.map((it,i)=>`${it.name}→D${emptyRows[i]}`).join(', ')}`);
    }

    const updates = items.map(item => ({
      range: `'回収'!${colStr}${nameMap[item.name]}`,
      values: [[item.amount]],
    }));
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: { valueInputOption: 'RAW', data: updates },
    });
    console.log(`[v2/回収/${label2}] ${d}日 ${updates.length}件 書込`);
  };

  await writeKaishuu(creaGeppoId, 'CREA', creaKaishuu);
  await writeKaishuu(fuwaGeppoId, 'ふわもこ', fuwaKaishuu);

  // ── SBシート書込 ────────────────────────
  const kyuryoRes = await sheets.spreadsheets.values.get({
    spreadsheetId: nippo.id,
    range: "'給料UI'!A4:T40",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const kyuryoRows = kyuryoRes.data.values || [];
  const extractSB = (nameIdx, valueStart) => {
    const items = [];
    for (const r of kyuryoRows) {
      const name = (r[nameIdx] || '').toString().trim();
      if (!name) continue;
      if (name.includes('合計') || name.includes('▌')) continue;
      const vals = [];
      for (let i = 0; i < 8; i++) vals.push(r[valueStart + i] ?? '');
      items.push({ name, vals });
    }
    return items;
  };
  const creaSB = extractSB(0, 1);
  const fuwaSB = extractSB(11, 12);

  const writeSB = async (sheetId, label2, items) => {
    if (items.length === 0) return;
    const hRes = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId, range: "'SB'!A7:KZ7",
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const header = hRes.data.values?.[0] || [];
    const dayColIdx  = header.findIndex(h => String(h).trim() === `${d}日`);
    const oneColIdx  = header.findIndex(h => String(h).trim() === '1日');
    if (dayColIdx === -1 || oneColIdx === -1) { console.log(`[v2/SB/${label2}] 列検出失敗`); return; }

    const nameColIdx = oneColIdx - 1;
    const nameColStr = idxToCol(nameColIdx);
    const startColStr = idxToCol(dayColIdx);
    const endColStr   = idxToCol(dayColIdx + 7);

    const nRes = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId, range: `'SB'!${nameColStr}1:${nameColStr}100`,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const nameCol = nRes.data.values || [];

    const normalize = (s) => String(s || '')
      .replace(/【[^】]*】/g, '')
      .replace(/\[[^\]]*\]/g, '')
      .split(/[\s　]/)[0]
      .trim();

    const nameMap = {};
    nameCol.forEach((r, i) => {
      const n = normalize(r[0]);
      if (n && i >= 8) nameMap[n] = i + 1;
    });

    const updates = [];
    const notFound = [];
    for (const item of items) {
      const row = nameMap[normalize(item.name)];
      if (!row) { notFound.push(item.name); continue; }
      updates.push({
        range: `'SB'!${startColStr}${row}:${endColStr}${row}`,
        values: [item.vals],
      });
    }
    if (updates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: { valueInputOption: 'RAW', data: updates },
      });
    }
    if (notFound.length) console.log(`[v2/SB/${label2}] 未発見: ${notFound.join(',')}`);
    console.log(`[v2/SB/${label2}] ${updates.length}件 書込`);
  };

  await writeSB(creaGeppoId, 'CREA', creaSB);
  await writeSB(fuwaGeppoId, 'ふわもこ', fuwaSB);

  // ── AC18 給料+交通費合計
  const salaryRes = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: nippo.id,
    ranges: ["'給料UI'!C4:D25", "'給料UI'!N4:O25"],
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const sumRange = (rows) => (rows || []).reduce((s, r) => {
    return s + (Number(r[0]) || 0) + (Number(r[1]) || 0);
  }, 0);
  const creaSalary = sumRange(salaryRes.data.valueRanges[0].values);
  const fuwaSalary = sumRange(salaryRes.data.valueRanges[1].values);
  await sheets.spreadsheets.values.update({
    spreadsheetId: creaGeppoId,
    range: `'売上'!${col}18`,
    valueInputOption: 'RAW',
    requestBody: { values: [[creaSalary]] },
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: fuwaGeppoId,
    range: `'売上'!${col}18`,
    valueInputOption: 'RAW',
    requestBody: { values: [[fuwaSalary]] },
  });

  // ── AC23 経費合計（CREAのみ）
  const p24Res = await sheets.spreadsheets.values.get({
    spreadsheetId: nippo.id,
    range: "'現金管理'!P24",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const keihiTotal = Number(p24Res.data.values?.[0]?.[0]) || 0;
  await sheets.spreadsheets.values.update({
    spreadsheetId: creaGeppoId,
    range: `'売上'!${col}23`,
    valueInputOption: 'RAW',
    requestBody: { values: [[keihiTotal]] },
  });

  // ── 人件費（CREAのみ）: 現金管理U5:X10 → 月報人件費D列マッチで該当日列に時間書込
  const attendRes = await sheets.spreadsheets.values.get({
    spreadsheetId: nippo.id,
    range: "'現金管理'!U5:X10",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const attendRows = attendRes.data.values || [];
  const attendItems = attendRows
    .map(r => ({ name: (r[0] || '').toString().trim(), hours: r[3] }))
    .filter(it => it.name && it.hours !== '' && it.hours !== null && it.hours !== undefined);

  if (attendItems.length > 0) {
    const nameRes = await sheets.spreadsheets.values.get({
      spreadsheetId: creaGeppoId,
      range: "'人件費'!D1:D50",
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const nameCol = nameRes.data.values || [];
    const nameMap = {};
    nameCol.forEach((r, i) => {
      const n = (r[0] || '').toString().trim();
      if (n) nameMap[n] = i + 1;
    });

    const jinkenCol = dayToJinkenCol(d);
    const jinkenUpdates = [];
    const notFound = [];
    for (const it of attendItems) {
      const row = nameMap[it.name];
      if (!row) { notFound.push(it.name); continue; }
      jinkenUpdates.push({
        range: `'人件費'!${jinkenCol}${row}`,
        values: [[it.hours]],
      });
    }
    if (jinkenUpdates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: creaGeppoId,
        requestBody: { valueInputOption: 'RAW', data: jinkenUpdates },
      });
    }
    if (notFound.length) console.log(`[v2/人件費/CREA] 名前未発見: ${notFound.join(',')}`);
    console.log(`[v2/人件費/CREA] ${d}日(${jinkenCol}列) ${jinkenUpdates.length}件 書込`);
  } else {
    console.log('[v2/人件費/CREA] 該当データなし');
  }

  console.log(`[v2] ${label} 完了`);

  // line_handler.js 互換の返り値
  return {
    label,
    totalSales: creaSales,
    total_hon:  creaHon,
    am_hon:     creaHonBySlot.am,
    pm_hon:     creaHonBySlot.pm,
    night_hon:  creaHonBySlot.night,
    total_count: creaShukkin,
    am_count:   creaU.am,
    pm_count:   creaU.pm,
    night_count: creaU.night,
    fuwamoko: {
      totalSales: fuwaSales,
      total_hon:  fuwaHon,
      am_hon:     fuwaHonBySlot.am,
      pm_hon:     fuwaHonBySlot.pm,
      night_hon:  fuwaHonBySlot.night,
      total_count: fuwaShukkin,
      am_count:   fuwaU.am,
      pm_count:   fuwaU.pm,
      night_count: fuwaU.night,
    },
  };
}

module.exports = { syncNippoToGeppo, dayToCol, findNippoByDate, findGeppoByStoreMonth };
