// shift_route_date.js - C-040 LINE「(M/D)の出勤更新して」コマンドルーター
// 作成: 2026-04-30 すい(Claude代行) v1.0
// 目的: 月またぎ等で漏れた日付シフトをLINEから手動で再発射可能にする
//
// 動作: シフトグループ内で「5/6の出勤更新して」のような投稿を検知
//   → 該当月SSから対象日の出勤者を抽出
//   → venrey_dispatch_specific.sh にカンマ区切りで発火
//   → LINE返信で結果通知
//
// ★編集権限: すいさん専属管轄 ★
// MONTHLY_SHEETS は main_test.py の同名辞書と同期させること。

const { exec } = require('child_process');
const { google } = require('googleapis');

const SHIFT_GROUP_ID = 'C4befe4675d94c864734eae6b897f1484';

const MONTHLY_SHEETS = {
  '2026-4': '1y33e0FlbS2R9d-iMQqaSm29rtYtl1JIcR1PpERwG2W4',
  '2026-5': '1XTmmkZP6k6PIhZ7RPHD75ClC_gfFA11U960wvCvEuB0',
  '2026-6': '1yizaAM_aQFaepv0kYlKBZY7uDUDsJA_9rwSn7hrDqKQ',
  '2026-7': '1k4FVGkUkqR1HUaroC0bH-REaAy8pEM3uFjO3b202e9o',
  '2026-8': '1en-dlxUDLmSEmSPp2mCBgYDYfDfmfKCf7GlPFAN8GSg',
  '2026-9': '1UeVn4llROiPOnISZ86KN3hXp0f4_RCARL07iEflBDNI',
  '2026-10': '19FbRRmAhQONKhsTK3jeHLJV_WLTcpmiP1n64ykAXrOM',
  '2026-11': '1XalqXt1sfswmPasUY2jMtTXduA6VR7JE-ieBDFKX06I',
  '2026-12': '1tuTfLTt1D1yXp8e86IICYkf4iAo9yHlJpr1OVNfZcQY',
};

let _auth = null;
function getAuth() {
  if (_auth) return _auth;
  _auth = new google.auth.GoogleAuth({
    keyFile: '/root/.config/chatbot-service-account.json',
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return _auth;
}

function isShiftValue(v) {
  if (v == null) return false;
  const s = String(v).trim();
  if (!s) return false;
  if (s === 'OFF' || s === '休' || s === '×' || s === '休み' || s === '-') return false;
  if (/^\d+$/.test(s)) return false;
  return true;
}

function isAggregateName(name) {
  return /^[📊🚨📅🏪]/.test(name);
}

function selectSheetId(month) {
  const now = new Date();
  let year = now.getFullYear();
  if (month < (now.getMonth() + 1)) year += 1;
  const key = `${year}-${month}`;
  return { ssId: MONTHLY_SHEETS[key], year, key };
}

async function getStaffOnDate(month, day) {
  const { ssId, year, key } = selectSheetId(month);
  if (!ssId) {
    return { error: `MONTHLY_SHEETS未登録: ${key}` };
  }
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const meta = await sheets.spreadsheets.get({ spreadsheetId: ssId });

  const dayStr = String(day);
  const result = { ssId, year, month, day, byStore: {}, totalCount: 0 };

  for (const sheet of meta.data.sheets) {
    const title = sheet.properties.title;
    if (title.includes('スタッフ管理')) continue;

    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: ssId,
      range: `'${title}'!A1:AZ80`,
    });
    const rows = r.data.values || [];
    if (rows.length === 0) continue;

    let headerRow = -1, targetCol = -1;
    for (let i = 0; i < Math.min(rows.length, 6); i++) {
      const row = rows[i] || [];
      for (let j = 0; j < row.length; j++) {
        const v = String(row[j] || '');
        if (new RegExp(`^${dayStr}[\\s\\n]*[月火水木金土日]`).test(v)) {
          headerRow = i;
          targetCol = j;
          break;
        }
      }
      if (headerRow >= 0) break;
    }
    if (headerRow < 0) {
      result.byStore[title] = { error: `${month}/${day}列が見つかりません` };
      continue;
    }
    const list = [];
    for (let i = headerRow + 1; i < rows.length; i++) {
      const name = String(rows[i][0] || '').trim();
      const val = String(rows[i][targetCol] || '').trim();
      if (!name || isAggregateName(name) || !isShiftValue(val)) continue;
      list.push({ name, shift: val });
    }
    result.byStore[title] = { list };
    result.totalCount += list.length;
  }
  return result;
}

async function handle(event, text, client) {
  if (event?.source?.type !== 'group' || event?.source?.groupId !== SHIFT_GROUP_ID) {
    return false;
  }
  const t = String(text || '').trim();
  // パターン: "5/6の出勤更新して" "5/6 出勤あげて" "5月6日の出勤更新" 等
  const m = t.match(/^(\d{1,2})[\/／月](\d{1,2})日?[\sのを　]*出勤[\sのを　]*(更新|あげ|挙げ|上げ|反映|送信)/);
  if (!m) return false;

  const month = parseInt(m[1], 10);
  const day = parseInt(m[2], 10);
  if (month < 1 || month > 12 || day < 1 || day > 31) return false;

  console.log(`[ShiftDate] 受信: ${month}/${day} 出勤更新指示`);

  setImmediate(async () => {
    try {
      const result = await getStaffOnDate(month, day);
      if (result.error) {
        await client.pushMessage({
          to: SHIFT_GROUP_ID,
          messages: [{ type: 'text', text: `❌ エラー: ${result.error}` }],
        }).catch(e => console.error('[ShiftDate] push err:', e.message));
        return;
      }

      const allNames = [];
      const summaryLines = [];
      for (const [store, data] of Object.entries(result.byStore)) {
        if (data.error) {
          summaryLines.push(`【${store}】${data.error}`);
          continue;
        }
        const names = data.list.map(x => x.name);
        if (names.length === 0) {
          summaryLines.push(`【${store}】出勤なし`);
          continue;
        }
        allNames.push(...names);
        const lines = data.list.map(x => `  ${x.name}: ${x.shift}`);
        summaryLines.push(`【${store}】${names.length}名\n${lines.join('\n')}`);
      }

      if (allNames.length === 0) {
        await client.pushMessage({
          to: SHIFT_GROUP_ID,
          messages: [{ type: 'text', text: `⚠️ ${month}/${day} の出勤者が見つかりません\n\n${summaryLines.join('\n\n')}` }],
        }).catch(e => console.error('[ShiftDate] push err:', e.message));
        return;
      }

      const staffArg = allNames.join(',').replace(/["$`\n\r]/g, '');
      const cmd = `/root/venrey_dispatch_specific.sh "${staffArg}"`;
      console.log(`[ShiftDate] dispatch ${allNames.length}名`);
      exec(cmd, (err, stdout, stderr) => {
        if (err) {
          console.error(`[ShiftDate] dispatch error: ${err.message}`);
          client.pushMessage({
            to: SHIFT_GROUP_ID,
            messages: [{ type: 'text', text: `❌ ベンリー発射エラー: ${err.message}` }],
          }).catch(() => {});
        } else {
          console.log(`[ShiftDate] dispatch OK: ${stdout.trim()}`);
        }
      });

      const replyText = `✅ ${month}/${day} の出勤を更新します（合計 ${allNames.length}名）\n\n${summaryLines.join('\n\n')}\n\nベンリー反映まで5〜10分かかります。`;
      await client.pushMessage({
        to: SHIFT_GROUP_ID,
        messages: [{ type: 'text', text: replyText }],
      }).catch(e => console.error('[ShiftDate] push err:', e.message));
    } catch (e) {
      console.error('[ShiftDate] handle error:', e.message, e.stack);
      try {
        await client.pushMessage({
          to: SHIFT_GROUP_ID,
          messages: [{ type: 'text', text: `❌ 内部エラー: ${e.message}` }],
        });
      } catch (_) {}
    }
  });

  return true;
}

// v1.1 (2026-05-08): isBotCommand 用判定エクスポート
// 1パーツの「(M/D)の出勤更新して」コマンドを line_handler.js の isBotCommand で許可するため
function isShiftDateCommand(text) {
  if (!text) return false;
  return /^\s*\d{1,2}[\/／月]\d{1,2}日?[\sのを　]*出勤[\sのを　]*(更新|あげ|挙げ|上げ|反映|送信)/.test(String(text).trim());
}

module.exports = { handle, SHIFT_GROUP_ID, getStaffOnDate, isShiftDateCommand };
