/**
 * kokyaku_update.js
 * 「○月○日 顧客更新」「【○月○日　顧客更新】」
 *  → 該当日の日報スプシの「日報_全件」B3:J を 利用履歴(1A_LaiWm…) の最初の空白行から追記
 * ※C-038: U列(キャンセル)=TRUEで備考先頭に【キャンセル】を付加
 */
const { google } = require('googleapis');
const { createAuthClient, findSpreadsheetByDateStr, normalizeDateStr } = require('./lib/common');
const customerSearch = require('./customer_search');

const KOKYAKU_TARGET_SHEET_ID   = '1A_LaiWm2QhvXk4jKINSRvKKX3-xKYLzygNnlY3ZtI9A';
const KOKYAKU_TARGET_SHEET_NAME = '利用履歴';
const NIPPO_ZENKEN_SHEET_NAME   = '日報_全件';

function parseKokyakuUpdateCommand(text) {
  const m = text.match(/^(\d{1,2}月\d{1,2}日)[\s　]+顧客更新$/);
  if (m) return { dateStr: m[1] };
  const m2 = text.match(/^【(\d{1,2}月\d{1,2}日)[\s　]+顧客更新】$/);
  if (m2) return { dateStr: m2[1] };
  return null;
}

function toSlashDate(nenGappiStr) {
  const m = nenGappiStr.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (!m) return nenGappiStr;
  return `${m[1]}/${parseInt(m[2])}/${parseInt(m[3])}`;
}

async function processKokyakuUpdate(dateStr) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const normalizedDate = normalizeDateStr(dateStr);
  const slashDate = toSlashDate(normalizedDate);
  const spreadsheetId = await findSpreadsheetByDateStr(dateStr);
  console.log(`[顧客更新] 日報ファイル: ${spreadsheetId}`);

  // 日報_全件 B3:U を grid データで取得（U列のチェックボックス boolValue を読むため）
  const srcRes = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [`'${NIPPO_ZENKEN_SHEET_NAME}'!B3:U`],
    includeGridData: true,
  });
  const gridData = srcRes.data.sheets?.[0]?.data?.[0];
  const rows = gridData?.rowData || [];

  const CANCEL_PREFIX = '【キャンセル】';
  const srcRows = [];
  for (const row of rows) {
    const cells = row.values || [];
    const formatted = cells.slice(0, 9).map(c => c?.formattedValue ?? ''); // B〜J = 9列
    const hasContent = formatted.some(v => v !== null && v !== undefined && String(v).trim() !== '');
    if (!hasContent) continue;
    // U列 = idx 19 (B=0, C=1, ..., U=19)
    const isCancel = cells[19]?.effectiveValue?.boolValue === true;
    if (isCancel) {
      const memo = formatted[5] || ''; // G列 = idx 5（B=0,C=1,D=2,E=3,F=4,G=5）
      formatted[5] = memo.startsWith(CANCEL_PREFIX) ? memo : CANCEL_PREFIX + memo;
    }
    srcRows.push(formatted);
  }
  if (srcRows.length === 0) return { count: 0, dateLabel: slashDate };

  // 利用履歴の最初の空白行を特定
  const targetRes = await sheets.spreadsheets.values.get({
    spreadsheetId: KOKYAKU_TARGET_SHEET_ID,
    range: `'${KOKYAKU_TARGET_SHEET_NAME}'!A:A`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const targetACol = targetRes.data.values || [];
  let firstEmptyRow = targetACol.length + 1;
  for (let i = targetACol.length - 1; i >= 0; i--) {
    if (targetACol[i] && targetACol[i][0] && String(targetACol[i][0]).trim() !== '') {
      firstEmptyRow = i + 2;
      break;
    }
  }
  console.log(`[顧客更新] 貼付け開始行: ${firstEmptyRow}`);

  const dateValues = srcRows.map(() => [slashDate]);
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: KOKYAKU_TARGET_SHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: `'${KOKYAKU_TARGET_SHEET_NAME}'!A${firstEmptyRow}`, values: dateValues },
        { range: `'${KOKYAKU_TARGET_SHEET_NAME}'!B${firstEmptyRow}`, values: srcRows },
      ],
    },
  });
  console.log(`[顧客更新] 書込完了: ${srcRows.length}行`);

  // 利用履歴キャッシュをクリア
  customerSearch.invalidateCache();

  return { count: srcRows.length, dateLabel: slashDate };
}

async function handleEvent(event, text, client) {
  const cmd = parseKokyakuUpdateCommand(text);
  if (!cmd) return false;
  const userId = event.source?.userId;
  const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
  const isGroup = event.source?.type === 'group';
  console.log(`[LINE] 顧客更新${isGroup ? '(グループ)' : ''}: ${cmd.dateStr}`);
  setImmediate(async () => {
    try {
      const result = await processKokyakuUpdate(cmd.dateStr);
      console.log(`[顧客更新] 完了: ${result.dateLabel} ${result.count}件`);
    } catch (err) {
      console.error('[顧客更新] エラー:', err.message);
      await client.pushMessage({
        to: target,
        messages: [{ type: 'text', text: `顧客更新エラー: ${err.message}` }],
      }).catch(() => {});
    }
  });
  return true;
}

module.exports = { handleEvent, parseKokyakuUpdateCommand, processKokyakuUpdate };
