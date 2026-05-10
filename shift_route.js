// shift_route.js - C-036 LINE→新SS シフト反映ルーター
// 作成: 2026-04-29 すい(Claude代行) - line_handler.js から物理分離
// 更新: 2026-04-29 すい(Claude代行) v5.4 - C-040連携追加 (LINE投稿→ベンリー強制反映)
// 目的: かずさん含む他者が line_handler.js を上書きしても、C-036本体ロジック(parts.length緩和を含む)は無傷
// このファイルは line_handler.js から require され、シフトグループ投稿を shift_reflect に委譲する
//
// ★編集権限: すいさん専属管轄 ★
// 他のスタッフは触らないこと。仕様変更が必要な場合は shift_reflect.js 側で対応。
// このファイルが破壊された場合は、サーバー上のガーディアン (cron 1分毎) が自動復旧します。

const SHIFT_GROUP_ID = 'C4befe4675d94c864734eae6b897f1484';

/**
 * シフトグループ投稿を shift_reflect に委譲する
 *
 * @param {object} event LINE webhook event
 * @param {string} text webhook テキスト
 * @param {object} client LINE messaging client (push用)
 * @returns {Promise<boolean>} true=このルートで処理した(line_handler は return すべき) / false=続行
 */
async function handle(event, text, client) {
  if (event?.source?.type !== 'group' || event?.source?.groupId !== SHIFT_GROUP_ID) {
    return false;
  }
  const parts = String(text || '').trim().split(/[\s　]+/);
  // すいさん v5.1 緩和: parts.length >= 2 で反応 (旧 >= 3 から緩和)
  if (parts.length < 2 && text !== '確認' && text !== '出勤確認') {
    return false;
  }
  setImmediate(async () => {
    try {
      const shiftReflect = require('./shift_reflect');
      const result = await shiftReflect.reflectShiftMessage(text);
      if (result.type === 'ignore') {
        console.log('[Shift] 非シフトメッセージ→スルー');
      } else if (result.type === 'success') {
        console.log(`[Shift] 結果: success ${result.writtenCount} 件 (${result.staffName} / ${result.store})`);
        // C-040 連携 (v5.4 追加): SS 書込成功スタッフを selected_staff として workflow_dispatch に渡し、
        // 既存値と同じでもベンリーに強制反映させる (main_test.py 側は SELECTED_STAFF 手動指定なら state cache 無視)。
        try {
          const sanitized = String(result.staffName || '').replace(/["\$`\n\r]/g, '');
          if (sanitized) {
            require('child_process').exec(
              `/root/venrey_dispatch_specific.sh "${sanitized}"`,
              (err) => { if (err) console.error('[Shift→Venrey] dispatch err:', err.message); }
            );
          }
        } catch (e) {
          console.error('[Shift→Venrey] exec failed:', e.message);
        }
      } else {
        console.log(`[Shift] 結果: ${JSON.stringify(result)}`);
      }
      const reply = shiftReflect.formatReply(result);
      if (reply) {
        await client.pushMessage({
          to: event.source.groupId,
          messages: [{ type: 'text', text: reply }],
        }).catch(() => {});
      }
    } catch (err) {
      console.error('[Shift] エラー:', err.message);
    }
  });
  return true;
}

module.exports = { handle, SHIFT_GROUP_ID };
