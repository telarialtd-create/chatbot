/**
 * meisai_handler.js
 * 明細関連:
 *  - 【日付,名前】「日付 名前」 → 旧明細(H2:M25) PNG
 *  - 【明細】複数行 → 明細書(A1:G47) PNG ※C-033
 */
const { google } = require('googleapis');
const {
  createAuthClient,
  findSpreadsheetByDateStr,
  normalizeSlashDate,
  parseAmount,
} = require('./lib/common');
const { screenshotMeisai } = require('./lib/canvas_render');

// ── 旧明細 ────────────────────────────────────────────────
const MEISAI_SHEET_NAME = '明細';
const MEISAI_RANGE = 'H2:M25';
const MEISAI_NAME_CELL = 'H5';

// ── 明細書（新フォーマット C-033）────────────────────────
const MEISAISYO_SHEET_NAME = '明細書';
const MEISAISYO_RANGE = 'A1:G47';
const MEISAISYO_NAME_CELL = 'C3';
const MEISAISYO_TRANSPORT_CELL = 'D40';
const MEISAISYO_BANCE_CELL = 'E41';
const MEISAISYO_OTSURI_CELL = 'E46';
const MEISAISYO_CHECK_CELL = 'E48';

const MEISAISYO_TRIGGER = /[【\[]\s*明\s*細\s*[】\]]/;

function normalizeTransportType(raw) {
  if (!raw) return null;
  const s = raw.trim();
  if (s === '片道') return '片道';
  if (s === '往復') return '往復';
  if (s.toUpperCase() === 'P' || s === 'ｐ') return 'P';
  if (s === 'なし' || s === 'ナシ') return 'なし';
  return null;
}

// 旧明細リクエストをパース
function parseMeisaiRequest(text) {
  if (/日時/.test(text) && /名前/.test(text)) {
    const dateMatch      = text.match(/日時[\s　:：]*(.+)/);
    const nameMatch      = text.match(/名前[\s　:：]*(.+)/);
    const transportMatch = text.match(/交通費[\s　:：]*(.+)/);
    if (dateMatch && nameMatch) {
      const transportRaw = transportMatch ? transportMatch[1].trim() : null;
      const transportType = normalizeTransportType(transportRaw);
      return { dateStr: dateMatch[1].trim(), name: nameMatch[1].trim(), transportType };
    }
  }
  const bracketMatch = text.match(/【(.+?)[,、](.+?)】/);
  if (bracketMatch) return { dateStr: bracketMatch[1].trim(), name: bracketMatch[2].trim(), transportType: null };
  const spaceMatch = text.match(/^((?:\d{4}年)?\d{1,2}月\d{1,2}日)[\s　]+(\S+)$/);
  if (spaceMatch) return { dateStr: spaceMatch[1].trim(), name: spaceMatch[2].trim(), transportType: null };
  return null;
}

// 明細書リクエストをパース（複数行）
function parseMeisaisyoRequest(text) {
  if (!text) return null;
  if (!MEISAISYO_TRIGGER.test(text)) return null;
  let dateStr = null, name = null, transport = null, bance = null, otsuri = null;
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
  return { dateStr, name, transport, bance, otsuri };
}

async function writeCellValue(spreadsheetId, sheetName, cell, value) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${sheetName}'!${cell}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[value]] },
  });
  console.log(`[明細] ${sheetName}!${cell} 書込: ${value}`);
}

// setMeisaiFromUriage 相当を Node で実装
async function runAppsScript(spreadsheetId) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

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

  const transportRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${MEISAI_SHEET_NAME}'!H18`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const transportType = ((transportRes.data.values || [[]])[0] || [])[0] || '';

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
    console.log('[明細] マスター取得エラー（交通費=0）:', e.message);
  }
  const finalTransportFee =
    transportType === '片道' ? baseTransportFee :
    transportType === '往復' ? baseTransportFee * 2 : 0;

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

  const urRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'売上'!AN:BB",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const urRows = (urRes.data.values || []).slice(1);
  const toK = v => (v === '' || v === null || v === undefined || isNaN(Number(v))) ? '' : Number(v) * 1000;
  const meisaiRows = [];
  let cashlessTotal = 0;

  for (const row of urRows) {
    if (row[2] !== genjimei) continue;
    const honsu   = row[0]  ?? '';
    const course  = row[8]  ?? '';
    const shimei  = row[3]  ?? '';
    const ryokin  = toK(row[12]);
    const kyuryo  = toK(row[14]);
    const option  = row[4]  ?? '';
    meisaiRows.push([honsu, course, shimei, ryokin, kyuryo, option]);
    if (/カード|ペイペイ|paypay/i.test(String(option)) && ryokin !== '') {
      cashlessTotal += Number(ryokin);
    }
  }

  const batchData = [];
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H8:M15`, values: Array(8).fill(Array(6).fill('')) });
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H36:M55`, values: Array(20).fill(Array(6).fill('')) });
  if (meisaiRows.length > 0) {
    batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H8`, values: meisaiRows });
    batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H36`, values: meisaiRows });
  }
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!H21`, values: [[finalTransportFee]] });
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!K21`, values: [[zenkaiValue]] });
  batchData.push({ range: `'${MEISAI_SHEET_NAME}'!L21`, values: [[cashlessTotal > 0 ? cashlessTotal : '']] });

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data: batchData },
  });
  console.log('[明細] setMeisaiFromUriage 完了');
}

// 旧明細処理
async function processMeisaiAndPush(target, client, dateStr, name, transportType = null) {
  try {
    const spreadsheetId = await findSpreadsheetByDateStr(dateStr);
    await writeCellValue(spreadsheetId, MEISAI_SHEET_NAME, MEISAI_NAME_CELL, name);
    if (transportType !== null) {
      await writeCellValue(spreadsheetId, MEISAI_SHEET_NAME, 'H18', transportType);
    }
    await runAppsScript(spreadsheetId);
    await new Promise(r => setTimeout(r, 3000));
    const filename = await screenshotMeisai(spreadsheetId, MEISAI_SHEET_NAME, MEISAI_RANGE, 'meisai');
    const baseUrl = (process.env.LINE_BOT_SERVER_URL || '').replace(/\/$/, '');
    const imageUrl = `${baseUrl}/temp/${filename}`;
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

// 明細書処理（C-033）
async function processMeisaisyoAndPush(target, client, parsed) {
  try {
    const spreadsheetId = await findSpreadsheetByDateStr(parsed.dateStr);
    const auth = createAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });

    // Phase1: 名前書込＋E41/E46クリア
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

    // Phase2: 残フィールド書込
    const data = [];
    if (parsed.transport != null) data.push({ range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_TRANSPORT_CELL}`, values: [[parsed.transport]] });
    if (parsed.bance != null)     data.push({ range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_BANCE_CELL}`, values: [[parsed.bance]] });
    if (parsed.otsuri != null)    data.push({ range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_OTSURI_CELL}`, values: [[parsed.otsuri]] });
    data.push({ range: `'${MEISAISYO_SHEET_NAME}'!${MEISAISYO_CHECK_CELL}`, values: [[true]] });

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data },
    });

    const filename = await screenshotMeisai(spreadsheetId, MEISAISYO_SHEET_NAME, MEISAISYO_RANGE, 'meisaisyo', { compactEmpty: true });
    const baseUrl = (process.env.LINE_BOT_SERVER_URL || '').replace(/\/$/, '');
    const imageUrl = `${baseUrl}/temp/${filename}`;
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

// LINEイベント処理（明細書→旧明細の順で判定）
async function handleEvent(event, text, client) {
  const userId = event.source?.userId;

  // 【明細】形式（明細書 C-033）
  const meisaisyo = parseMeisaisyoRequest(text);
  if (meisaisyo) {
    const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
    console.log(`[LINE] [C-033] 明細書受信: ${meisaisyo.dateStr}/${meisaisyo.name}/交=${meisaisyo.transport ?? '-'}/B=${meisaisyo.bance ?? '-'}/O=${meisaisyo.otsuri ?? '-'} → ${target}`);
    setImmediate(() => processMeisaisyoAndPush(target, client, meisaisyo).catch(err => console.error('[明細書] 未処理:', err.message)));
    return true;
  }

  // 【日付,名前】形式（旧明細）
  const meisai = parseMeisaiRequest(text);
  if (meisai) {
    const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
    console.log(`[LINE] 明細受信: ${meisai.dateStr}/${meisai.name}/交=${meisai.transportType ?? '-'} → ${target}`);
    setImmediate(() => processMeisaiAndPush(target, client, meisai.dateStr, meisai.name, meisai.transportType).catch(err => console.error('[明細] 未処理:', err.message)));
    return true;
  }

  return false;
}

module.exports = { handleEvent, parseMeisaiRequest, parseMeisaisyoRequest, processMeisaiAndPush, processMeisaisyoAndPush };
