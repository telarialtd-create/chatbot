/**
 * geppo_update.js
 * 「月報更新」「月報更新 4月7日」「4月7日 月報更新」
 *  → nippo_to_geppo_v2.syncNippoToGeppo を起動して結果通知
 *
 * ⚠️ 現状このファイルは require されていないデッドコード（2026-05-19時点）
 *    実際の月報更新は line_handler.js 内のハンドラで完結。
 *    将来再有効化する場合は handleEvent に storeId/folderInfo を渡す改修必須。
 */
const { syncNippoToGeppo } = require('./nippo_to_geppo_v2');
// T-001 専用 folderInfo（このラッパーは現状T-001専用前提）
const T001_FOLDER_INFO = {
  storeId: 'T-001',
  storeName: 'CREA',
  folderId: process.env.DATA_FOLDER_ID || '16R1BK5NnvYkH4Eqh6t51OGQl0tXVJ3If',
};

// C-029 v5 2026-05-19: storeLabel 動的化 + ふわもこ無しなら2店舗合算省略
function buildResultMessage(result) {
  const lines = [
    `✅ 月報更新完了`,
    `📅 ${result.label}`,
    ``,
    `━━ ${result.storeLabel} ━━`,
    `💰 総売上: ${result.totalSales.toLocaleString()}円`,
    `📊 本数: ${result.total_hon}本（朝${result.am_hon}/昼${result.pm_hon}/夜${result.night_hon}）`,
    `👥 出勤: ${result.total_count}人（朝${result.am_count}/昼${result.pm_count}/夜${result.night_count}）`,
    `💵 現金残り: ${(result.genkinNokori || 0).toLocaleString()}円`,
  ];
  const fw = result.fuwamoko;
  if (fw) {
    lines.push(
      ``,
      `━━ ${fw.storeLabel} ━━`,
      `💰 総売上: ${fw.totalSales.toLocaleString()}円`,
      `📊 本数: ${fw.total_hon}本（朝${fw.am_hon}/昼${fw.pm_hon}/夜${fw.night_hon}）`,
      `👥 出勤: ${fw.total_count}人（朝${fw.am_count}/昼${fw.pm_count}/夜${fw.night_count}）`,
      `💵 現金残り: ${(fw.genkinNokori || 0).toLocaleString()}円`,
      ``,
      `━━ 2店舗合算 ━━`,
      `💰 総売上: ${(result.totalSales + fw.totalSales).toLocaleString()}円`,
      `📊 本数: ${result.total_hon + fw.total_hon}本`,
      `👥 出勤: ${result.total_count + fw.total_count}人`,
      `💵 現金残り: ${((result.genkinNokori || 0) + (fw.genkinNokori || 0)).toLocaleString()}円`,
    );
  }
  return lines.join('\n');
}

function buildWarningMessage(warnings) {
  const grouped = {};
  for (const w of warnings) (grouped[w.section] ||= []).push(w.name);
  const lines = Object.entries(grouped).map(([sec, names]) => `・${sec}: ${names.join(', ')}`).join('\n');
  return `⚠️ 空白行不足で書き込めなかった名前があります\n${lines}\n\n該当シートで空白行を追加するか、不要な行を整理してください。`;
}

async function processGeppoUpdate(target, client, text) {
  try {
    let syncDate = null;
    const dateMatch = text.match(/(\d{1,2})[月\/](\d{1,2})日?/);
    if (dateMatch) {
      const jst = new Date(Date.now() + 9 * 3600 * 1000);
      syncDate = new Date(Date.UTC(jst.getUTCFullYear(), parseInt(dateMatch[1]) - 1, parseInt(dateMatch[2])));
    }
    // C-029 2026-05-19: 新シグネチャ対応
    const result = await syncNippoToGeppo({ jstDate: syncDate, storeId: 'T-001', folderInfo: T001_FOLDER_INFO });
    await client.pushMessage({
      to: target,
      messages: [{ type: 'text', text: buildResultMessage(result) }],
    });
    if (result.warnings && result.warnings.length > 0) {
      await client.pushMessage({
        to: target,
        messages: [{ type: 'text', text: buildWarningMessage(result.warnings) }],
      }).catch(() => {});
    }
  } catch (err) {
    console.error('[月報] エラー:', err.message);
    await client.pushMessage({
      to: target,
      messages: [{ type: 'text', text: `❌ 月報更新エラー: ${err.message}` }],
    }).catch(() => {});
  }
}

async function handleEvent(event, text, client) {
  if (!text.includes('月報更新')) return false;
  const userId = event.source?.userId;
  const target = event.source?.groupId || userId || process.env.LINE_USER_ID;
  setImmediate(() => processGeppoUpdate(target, client, text));
  return true;
}

module.exports = { handleEvent, processGeppoUpdate };
