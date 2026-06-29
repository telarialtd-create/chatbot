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

// データフォルダ（日報・月報の両方が入っている）= T-001 自社案件専用
// ⚠️ 直接参照禁止。syncNippoToGeppo の folderInfo.folderId 経由で動的に解決すること。
// （2026-05-19 C-029: T-1043指定でT-001月報誤更新事故の再発防止）
const DATA_FOLDER_ID = process.env.DATA_FOLDER_ID || '16R1BK5NnvYkH4Eqh6t51OGQl0tXVJ3If';

// ── 月報シリーズ設定 ─────────────────────────
// 設計方針（2026-05-19 オーナー指示）:
//   T-001 (CREA) だけが「CREA + ふわもこ」の2系列という特殊パターン。
//   それ以外の全店舗は「店舗名そのままで1系列」のデフォルト挙動で動く。
//   → 新店舗追加時の設定変更は不要。月報シートを「{店舗名}売上YYYY-M月」で
//     店舗フォルダに置くだけで自動的に動く。
//   → 月報シートが存在しなければ findGeppoByStoreMonth が自然にエラーで止まる
//     ので、誤更新の防御は維持される（書込先が無ければ書き込めない）。
const STORE_GEPPO_SERIES_OVERRIDE = {
  'T-001': {
    crea:     { sheetPrefix: 'CREA',     storeFieldValue: 'CREA',     label: 'CREA'        },
    fuwamoko: { sheetPrefix: 'ふわもこ', storeFieldValue: 'ふわもこ', label: 'ふわもこSPA' },
  },
};

// 店舗の月報シリーズ設定を取得
// - T-001 はオーバーライド適用
// - その他は folderInfo.storeName を sheetPrefix / label にした1系列（CREA同等）
function resolveGeppoSeries(storeId, folderInfo) {
  const override = STORE_GEPPO_SERIES_OVERRIDE[storeId];
  if (override) return override;
  if (!folderInfo || !folderInfo.storeName) {
    throw new Error(`resolveGeppoSeries: ${storeId} の folderInfo.storeName が空です`);
  }
  return {
    crea: {
      sheetPrefix: folderInfo.storeName,  // 例: "Angel Spa" → 「Angel Spa売上YYYY-M月」を検索
      storeFieldValue: 'CREA',             // 日報の店舗フィールド判定値（テンプレ共通のCREA前提）
      label: folderInfo.storeName,         // LINE返信メッセージ用のラベル（"Angel Spa" 等）
    },
    // fuwamoko: なし
  };
}

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
// folderId は必須引数（C-029 2026-05-19: T-001fallback廃止）
// 検索順: ①完全一致「2026年5月19日」 → ②末尾一致「Angel Spa　2026年5月19日」等
async function findNippoByDate(jstDate, folderId) {
  if (!folderId) throw new Error('findNippoByDate: folderId は必須引数です');
  const y = jstDate.getUTCFullYear();
  const m = jstDate.getUTCMonth() + 1;
  const d = jstDate.getUTCDate();
  const dateStr = `${y}年${m}月${d}日`;

  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  // contains クエリで「日付文字列を含む」全ファイルを取得 → 完全一致 or 末尾一致で絞込
  const res = await drive.files.list({
    q: `'${folderId}' in parents and name contains '${dateStr}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: 'files(id,name)',
    pageSize: 20,
  });
  const files = res.data.files || [];
  // ① 完全一致を最優先
  const exact = files.find(f => f.name.trim() === dateStr);
  if (exact) return { id: exact.id, name: exact.name };
  // ② 末尾一致（「{店舗名}　{日付}」「{店舗名} {日付}」等のクライアント店舗命名に対応）
  const endsWith = files.find(f => f.name.trim().endsWith(dateStr));
  if (endsWith) return { id: endsWith.id, name: endsWith.name };
  // 見つからない
  throw new Error(
    `日報ファイルが見つかりません: ${dateStr} (folder=${folderId}` +
    (files.length ? ` / 候補=${files.map(f => f.name).join(',')}` : '') + ')'
  );
}

// folderId は必須引数（C-029 2026-05-19: T-001fallback廃止）
// 検索順: ①完全一致「Angel Spa売上2026-5月」 → ②末尾一致（前置店舗名等の許容）
async function findGeppoByStoreMonth(storeName, year, month, folderId) {
  if (!folderId) throw new Error('findGeppoByStoreMonth: folderId は必須引数です');
  const name = `${storeName}売上${year}-${month}月`;
  const monthTail = `売上${year}-${month}月`;  // 末尾一致用
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const res = await drive.files.list({
    q: `'${folderId}' in parents and name contains '${name}' and trashed=false`,
    fields: 'files(id,name,mimeType,shortcutDetails)',
    pageSize: 10,
  });
  const files = res.data.files || [];
  // ① 完全一致を最優先
  let target = files.find(f => f.name.trim() === name);
  // ② 末尾一致（「店舗名前置」等の命名揺れ吸収。ただし storeName 文字列を含むものに限定）
  if (!target) {
    target = files.find(f => {
      const n = f.name.trim();
      return n.endsWith(monthTail) && n.includes(storeName);
    });
  }
  if (!target) throw new Error(`月報シートが見つかりません: ${name} (folder=${folderId})`);
  // ショートカットなら実体ファイルIDに解決（2026-5月以降の運用変化に対応）
  if (target.mimeType === 'application/vnd.google-apps.shortcut' && target.shortcutDetails && target.shortcutDetails.targetId) {
    console.log(`[v2] "${target.name}" はショートカット → 実体 ${target.shortcutDetails.targetId} を使用`);
    return target.shortcutDetails.targetId;
  }
  return target.id;
}

// ── メイン同期関数 ───────────────────────────
// シグネチャ変更（C-029 2026-05-19）: { jstDate, storeId, folderInfo } の名前付き引数必須
// - storeId: 'T-001' などの店舗ID（必須）
// - folderInfo: { storeId, storeName, folderId } getFolderByStoreId() の戻り値（必須）
// - jstDate: 同期対象日（省略時は昨日）
//
// 旧シグネチャ syncNippoToGeppo(jstDate) は禁止。引数オブジェクト未指定なら即throw。
async function syncNippoToGeppo(opts) {
  // ── 第1層：シグネチャで必須引数を強制 ──
  if (!opts || typeof opts !== 'object' || opts instanceof Date) {
    throw new Error(
      'syncNippoToGeppo: 引数は { jstDate, storeId, folderInfo } 形式必須。' +
      '旧シグネチャ(jstDateのみ)は2026-05-19で廃止しました。'
    );
  }
  const { jstDate: inputDate, storeId, folderInfo } = opts;
  if (!storeId) throw new Error('syncNippoToGeppo: storeId は必須引数です');
  if (!folderInfo || !folderInfo.folderId) {
    throw new Error('syncNippoToGeppo: folderInfo.folderId は必須引数です');
  }
  if (folderInfo.storeId && folderInfo.storeId !== storeId) {
    throw new Error(
      `syncNippoToGeppo: storeId(${storeId}) と folderInfo.storeId(${folderInfo.storeId}) が不一致です`
    );
  }

  // ── 第2層：シリーズ設定を解決（T-001は2系列、それ以外は1系列の自動デフォルト）──
  // 月報シートが店舗フォルダに無ければ findGeppoByStoreMonth が自然にエラーで止まる
  // ので、ホワイトリストは不要（書込先が無ければ書き込めない）。

  let jstDate = inputDate;
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

  console.log(`[v2] ${label} 同期開始 (store=${storeId} folder=${folderInfo.folderId})`);

  // 店舗のシリーズ設定（T-001は2系列・それ以外は1系列の自動デフォルト）
  const series = resolveGeppoSeries(storeId, folderInfo);
  const hasFuwa = !!series.fuwamoko;
  const creaCfg = series.crea;
  const fuwaCfg = series.fuwamoko || null;

  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // 日報SS検索（folderInfo.folderId 配下に限定）
  const nippo = await findNippoByDate(jstDate, folderInfo.folderId);
  console.log(`[v2] 日報: ${nippo.name} (${nippo.id})`);

  // 月報SS検索（店舗×月・folderInfo.folderId 配下に限定）
  const creaGeppoId = await findGeppoByStoreMonth(creaCfg.sheetPrefix, y, m, folderInfo.folderId);
  const fuwaGeppoId = hasFuwa
    ? await findGeppoByStoreMonth(fuwaCfg.sheetPrefix, y, m, folderInfo.folderId)
    : null;
  console.log(`[v2] CREA系月報(${creaCfg.sheetPrefix})=${creaGeppoId}` +
    (hasFuwa ? `, FUWA系月報(${fuwaCfg.sheetPrefix})=${fuwaGeppoId}` : ', FUWA系=なし'));

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
    range: "'日報_全件'!B3:O1000",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const zenRows = zenRes.data.values || [];
  let creaHonBySlot = { am:0, pm:0, night:0 };
  let fuwaHonBySlot = { am:0, pm:0, night:0 };
  let creaCard = 0, fuwaCard = 0, creaPay = 0, fuwaPay = 0;
  for (const r of zenRows) {
    const store = (r[0] || '').toString().trim();
    // store フィールド値による系列判定（config の storeFieldValue を使用）
    // 単一系列店（Angel Spa等）は日報自体が1店舗専用のため全行カウント。
    // CREA(2系列)のみ store===CREA で系列分離する（2026-06-29 K-003 本数0バグ修正）。
    const isCrea = !hasFuwa || store === creaCfg.storeFieldValue;
    const isFuwa = hasFuwa && store.includes(fuwaCfg.storeFieldValue);
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
  if (hasFuwa) {
    await writeSales(fuwaGeppoId, 'ふわもこ', {
      sales: fuwaSales, am_hon: fuwaHonBySlot.am, pm_hon: fuwaHonBySlot.pm, night_hon: fuwaHonBySlot.night,
      hon: fuwaHon, am_u: fuwaU.am, pm_u: fuwaU.pm, night_u: fuwaU.night, shukkin: fuwaShukkin,
      taiki: fuwaTaiki, card: fuwaCard, pay: fuwaPay,
    });
  }

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
    ranges: ["'現金管理'!A5:H34", "'現金管理'!A37:H80"],
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
      const sum = fn + gn;
      if (sum === 0) continue;
      items.push({ name, amount: -sum });
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
    const overflow = [];
    if (toAdd.length > 0) {
      const writable = toAdd.slice(0, emptyRows.length);
      const overflowItems = toAdd.slice(emptyRows.length);
      overflowItems.forEach(it => overflow.push(it.name));

      if (writable.length > 0) {
        const addUpdates = writable.map((it, idx) => ({
          range: `'回収'!D${emptyRows[idx]}`,
          values: [[it.name]],
        }));
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: sheetId,
          requestBody: { valueInputOption: 'RAW', data: addUpdates },
        });
        writable.forEach((it, idx) => { nameMap[it.name] = emptyRows[idx]; });
        console.log(`[v2/回収/${label2}] 新規登録: ${writable.map((it,i)=>`${it.name}→D${emptyRows[i]}`).join(', ')}`);
      }
    }

    const updates = items
      .filter(item => nameMap[item.name])
      .map(item => ({
        range: `'回収'!${colStr}${nameMap[item.name]}`,
        values: [[item.amount]],
      }));
    if (updates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: { valueInputOption: 'RAW', data: updates },
      });
    }
    console.log(`[v2/回収/${label2}] ${d}日 ${updates.length}件 書込`);
    return overflow.map(name => ({ section: `回収/${label2}`, name }));
  };

  const warnings = [];
  warnings.push(...(await writeKaishuu(creaGeppoId, 'CREA', creaKaishuu)) || []);
  if (hasFuwa) {
    warnings.push(...(await writeKaishuu(fuwaGeppoId, 'ふわもこ', fuwaKaishuu)) || []);
  }

  // ── バンス書込（現金管理H列 → 回収シート「{名前}バンス」行 該当日列・H列値そのまま） ─
  const buildBansu = (rows) => {
    const items = [];
    for (const r of rows) {
      const name = (r[0] || '').toString().trim();
      if (!name || name.includes('▌')) continue;
      if (name.replace(/\s/g, '').includes('合計')) continue;
      const v = Number(r[7]);
      if (isNaN(v) || v === 0) continue;
      items.push({ name: `${name}バンス`, amount: v });
    }
    return items;
  };
  const creaBansu = buildBansu(creaRows);
  const fuwaBansu = buildBansu(fuwaRows);

  const writeBansu = async (sheetId, label2, items) => {
    if (items.length === 0) return [];
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: "'回収'!A1:AK200",
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const rows = res.data.values || [];
    const header = rows[0] || [];
    const colIdx = header.findIndex(h => String(h).trim() === `${d}日`);
    if (colIdx === -1) { console.log(`[v2/バンス/${label2}] ${d}日 列なし`); return []; }
    const colStr = idxToCol(colIdx);

    const MAX_ROW = 46;
    const nameMap = {};
    const emptyRows = [];
    for (let i = 1; i < rows.length; i++) {
      const row = i + 1;
      if (row > MAX_ROW) break;
      const v = (rows[i]?.[3] || '').toString().trim();
      if (v) nameMap[v] = row;
      else emptyRows.push(row);
    }

    const overflow = [];
    const addNameUpdates = [];
    for (const it of items) {
      if (nameMap[it.name]) continue;
      if (emptyRows.length === 0) {
        overflow.push(it.name);
        continue;
      }
      const row = emptyRows.shift();
      addNameUpdates.push({ range: `'回収'!D${row}`, values: [[it.name]] });
      nameMap[it.name] = row;
    }
    if (addNameUpdates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: { valueInputOption: 'RAW', data: addNameUpdates },
      });
      console.log(`[v2/バンス/${label2}] 新規登録 ${addNameUpdates.length}件`);
    }

    const updates = items
      .filter(it => nameMap[it.name])
      .map(it => ({ range: `'回収'!${colStr}${nameMap[it.name]}`, values: [[it.amount]] }));
    if (updates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: { valueInputOption: 'RAW', data: updates },
      });
    }
    console.log(`[v2/バンス/${label2}] ${d}日 ${updates.length}件 書込${overflow.length ? ` / 空白不足: ${overflow.join(',')}` : ''}`);
    return overflow.map(name => ({ section: `バンス/${label2}`, name }));
  };
  warnings.push(...(await writeBansu(creaGeppoId, 'CREA', creaBansu)) || []);
  if (hasFuwa) {
    warnings.push(...(await writeBansu(fuwaGeppoId, 'ふわもこ', fuwaBansu)) || []);
  }

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

  // 交通費は給料UIではなく「現金管理」D列を参照（既に kaishuuSrc で A5:G34/A37:G80 取得済み）
  const buildKotsuMap = (rows) => {
    const map = {};
    for (const r of rows) {
      const name = (r[0] || '').toString().trim();
      if (!name || name.includes('▌')) continue;
      if (name.replace(/\s/g, '').includes('合計')) continue;
      const k = Number(r[3]);
      if (!isNaN(k)) map[name] = k;
    }
    return map;
  };
  const creaKotsuMap = buildKotsuMap(creaRows);
  const fuwaKotsuMap = buildKotsuMap(fuwaRows);
  const overrideKotsu = (items, map) => {
    // vals = [売, 給料, 交通費, 指, フリー, リピ, X, 時間] → 交通費は idx 2
    for (const it of items) {
      if (map[it.name] !== undefined) it.vals[2] = map[it.name];
    }
  };
  overrideKotsu(creaSB, creaKotsuMap);
  overrideKotsu(fuwaSB, fuwaKotsuMap);

  const writeSB = async (sheetId, label2, items) => {
    if (items.length === 0) return [];
    const hRes = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId, range: "'SB'!A7:ZZ7",
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const header = hRes.data.values?.[0] || [];
    const dayColIdx  = header.findIndex(h => String(h).trim() === `${d}日`);
    const oneColIdx  = header.findIndex(h => String(h).trim() === '1日');
    if (dayColIdx === -1 || oneColIdx === -1) { console.log(`[v2/SB/${label2}] 列検出失敗`); return []; }

    // 名前列はA列固定（T列等は無視）
    const nameColStr = 'A';
    const startColStr = idxToCol(dayColIdx);
    const endColStr   = idxToCol(dayColIdx + 7);

    const nRes = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId, range: `'SB'!A1:A100`,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const nameCol = nRes.data.values || [];

    const normalize = (s) => String(s || '')
      .replace(/【[^】]*】/g, '')
      .replace(/\[[^\]]*\]/g, '')
      .split(/[\s　]/)[0]
      .trim();

    // 行9〜48 を対象範囲とする（i=8..47）
    const ROW_START = 9, ROW_END = 48;
    const nameMap = {};
    const emptyRows = [];
    for (let i = ROW_START - 1; i <= ROW_END - 1; i++) {
      const n = normalize(nameCol[i]?.[0]);
      const row = i + 1;
      if (n) nameMap[n] = row;
      else emptyRows.push(row);
    }

    const overflow = [];
    const addNameUpdates = [];
    for (const item of items) {
      if (nameMap[normalize(item.name)]) continue;
      if (emptyRows.length === 0) {
        overflow.push(item.name);
        continue;
      }
      const row = emptyRows.shift();
      addNameUpdates.push({ range: `'SB'!A${row}`, values: [[item.name]] });
      nameMap[normalize(item.name)] = row;
    }
    if (addNameUpdates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: { valueInputOption: 'RAW', data: addNameUpdates },
      });
      console.log(`[v2/SB/${label2}] 新規登録 ${addNameUpdates.length}件`);
    }

    const updates = [];
    for (const item of items) {
      const row = nameMap[normalize(item.name)];
      if (!row) continue;
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
    console.log(`[v2/SB/${label2}] ${updates.length}件 書込${overflow.length ? ` / 空白不足: ${overflow.join(',')}` : ''}`);
    return overflow.map(name => ({ section: `SB/${label2}`, name }));
  };

  warnings.push(...(await writeSB(creaGeppoId, 'CREA', creaSB)) || []);
  if (hasFuwa) {
    warnings.push(...(await writeSB(fuwaGeppoId, 'ふわもこ', fuwaSB)) || []);
  }

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
      range: "'人件費'!D1:D200",
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const nameCol = nameRes.data.values || [];
    const nameMap = {};
    const emptyRows = [];
    const SCAN_ROWS = 200;
    for (let i = 0; i < SCAN_ROWS; i++) {
      const n = (nameCol[i]?.[0] || '').toString().trim();
      const row = i + 1;
      if (n) nameMap[n] = row;
      else emptyRows.push(row);
    }

    const jinkenCol = dayToJinkenCol(d);
    const addNameUpdates = [];
    const overflow = [];
    for (const it of attendItems) {
      if (nameMap[it.name]) continue;
      if (emptyRows.length === 0) {
        overflow.push(it.name);
        continue;
      }
      const row = emptyRows.shift();
      addNameUpdates.push({ range: `'人件費'!D${row}`, values: [[it.name]] });
      nameMap[it.name] = row;
    }
    if (addNameUpdates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: creaGeppoId,
        requestBody: { valueInputOption: 'RAW', data: addNameUpdates },
      });
      console.log(`[v2/人件費/CREA] 新規登録 ${addNameUpdates.length}件`);
    }

    const jinkenUpdates = [];
    for (const it of attendItems) {
      const row = nameMap[it.name];
      if (!row) continue;
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
    console.log(`[v2/人件費/CREA] ${d}日(${jinkenCol}列) ${jinkenUpdates.length}件 書込${overflow.length ? ` / 空白不足: ${overflow.join(',')}` : ''}`);
    overflow.forEach(name => warnings.push({ section: '人件費/CREA', name }));
  } else {
    console.log('[v2/人件費/CREA] 該当データなし');
  }

  console.log(`[v2] ${label} 完了`);

  // ── 現金残り読み取り（月報「売上」タブ F列ラベル「現金残り」行・当日列） ──
  // C-029 2026-06-17: LINE返信メッセージに現金残りを追記（全店舗汎用・行番号非依存のラベル検索）
  const readGenkin = async (geppoId) => {
    if (!geppoId) return 0;
    try {
      const gr = await sheets.spreadsheets.values.batchGet({
        spreadsheetId: geppoId,
        ranges: ["'売上'!F1:F60", `'売上'!${col}1:${col}60`],
        valueRenderOption: 'UNFORMATTED_VALUE',
      });
      const labels = gr.data.valueRanges[0].values || [];
      const vals   = gr.data.valueRanges[1].values || [];
      for (let i = 0; i < labels.length; i++) {
        if ((labels[i]?.[0] || '').toString().trim() === '現金残り') {
          return Number(vals[i]?.[0]) || 0;
        }
      }
      console.log('[v2/現金残り] ラベル「現金残り」が見つかりませんでした (geppoId=' + geppoId + ')');
      return 0;
    } catch (e) {
      console.log('[v2/現金残り] 読み取り失敗: ' + e.message);
      return 0;
    }
  };
  const creaGenkin = await readGenkin(creaGeppoId);
  const fuwaGenkin = hasFuwa ? await readGenkin(fuwaGeppoId) : 0;

  // line_handler.js 互換の返り値
  return {
    label,
    warnings,
    storeLabel: creaCfg.label,   // ← LINE返信メッセージのCREA系列セクション名（"CREA" or "Angel Spa" 等）
    totalSales: creaSales,
    total_hon:  creaHon,
    am_hon:     creaHonBySlot.am,
    pm_hon:     creaHonBySlot.pm,
    night_hon:  creaHonBySlot.night,
    total_count: creaShukkin,
    am_count:   creaU.am,
    pm_count:   creaU.pm,
    night_count: creaU.night,
    genkinNokori: creaGenkin,
    // ふわもこ系列：無い店舗では null を返す（LINE返信側で省略判定）
    fuwamoko: hasFuwa ? {
      storeLabel: fuwaCfg.label,
      totalSales: fuwaSales,
      total_hon:  fuwaHon,
      am_hon:     fuwaHonBySlot.am,
      pm_hon:     fuwaHonBySlot.pm,
      night_hon:  fuwaHonBySlot.night,
      total_count: fuwaShukkin,
      am_count:   fuwaU.am,
      pm_count:   fuwaU.pm,
      night_count: fuwaU.night,
      genkinNokori: fuwaGenkin,
    } : null,
  };
}

module.exports = { syncNippoToGeppo, dayToCol, findNippoByDate, findGeppoByStoreMonth };
