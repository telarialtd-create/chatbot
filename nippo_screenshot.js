/**
 * nippo_screenshot.js
 * 「教えて」→ 当日の日報 CA3:CR32 を PNG にして送信
 */
const {
  findSpreadsheetId,
  getSheetBusinessDate,
  dateToNippoName,
  getPushTarget,
} = require('./lib/common');
const { screenshotCells } = require('./lib/canvas_render');

const TARGET_GID = 1873674341;
const RANGE = 'CA3:CR32';

async function processAndPush(target, client) {
  try {
    const date = getSheetBusinessDate();
    console.log(`[教えて] 検索日付(JST): ${dateToNippoName(date)}`);
    const spreadsheetId = await findSpreadsheetId(date);
    const filename = await screenshotCells(spreadsheetId, TARGET_GID, RANGE);

    const baseUrl = (process.env.LINE_BOT_SERVER_URL || '').replace(/\/$/, '');
    const imageUrl = `${baseUrl}/temp/${filename}`;
    console.log(`[教えて] Push送信: ${imageUrl} → ${target}`);
    await client.pushMessage({
      to: target,
      messages: [{ type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl }],
    });
  } catch (err) {
    console.error('[教えて] Push エラー:', err.message);
    await client.pushMessage({
      to: target,
      messages: [{ type: 'text', text: `エラー: ${err.message}` }],
    }).catch(() => {});
  }
}

// LINEイベントを処理（「教えて」を含めば true）
async function handleEvent(event, text, client) {
  if (!text.includes('教えて')) return false;
  const target = getPushTarget(event);
  console.log(`[LINE] (教えて) 受信 → バックグラウンド処理開始 target=${target}`);
  setImmediate(() => processAndPush(target, client).catch(err => console.error('[教えて] 未処理エラー:', err.message)));
  return true;
}

module.exports = { handleEvent, processAndPush };
