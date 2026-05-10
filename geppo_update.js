/**
 * geppo_update.js
 * 「月報更新」「月報更新 4月7日」「4月7日 月報更新」
 *  → nippo_to_geppo_v2.syncNippoToGeppo を起動して結果通知
 */
const { syncNippoToGeppo } = require('./nippo_to_geppo_v2');

function buildResultMessage(result) {
  const fw = result.fuwamoko;
  const combinedSales = result.totalSales + fw.totalSales;
  const combinedHon   = result.total_hon  + fw.total_hon;
  const combinedCount = result.total_count + fw.total_count;
  return (
    `✅ 月報更新完了\n` +
    `📅 ${result.label}\n\n` +
    `━━ CREA ━━\n` +
    `💰 総売上: ${result.totalSales.toLocaleString()}円\n` +
    `📊 本数: ${result.total_hon}本（朝${result.am_hon}/昼${result.pm_hon}/夜${result.night_hon}）\n` +
    `👥 出勤: ${result.total_count}人（朝${result.am_count}/昼${result.pm_count}/夜${result.night_count}）\n\n` +
    `━━ ふわもこSPA ━━\n` +
    `💰 総売上: ${fw.totalSales.toLocaleString()}円\n` +
    `📊 本数: ${fw.total_hon}本（朝${fw.am_hon}/昼${fw.pm_hon}/夜${fw.night_hon}）\n` +
    `👥 出勤: ${fw.total_count}人（朝${fw.am_count}/昼${fw.pm_count}/夜${fw.night_count}）\n\n` +
    `━━ 2店舗合算 ━━\n` +
    `💰 総売上: ${combinedSales.toLocaleString()}円\n` +
    `📊 本数: ${combinedHon}本\n` +
    `👥 出勤: ${combinedCount}人`
  );
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
    const result = await syncNippoToGeppo(syncDate);
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
