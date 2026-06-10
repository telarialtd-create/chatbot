// shift_route.js - C-036 LINE→新SS シフト反映ルーター
// 作成: 2026-04-29 すい(Claude代行) - line_handler.js から物理分離
// 更新: 2026-04-29 すい(Claude代行) v5.4 - C-040連携追加 (LINE投稿→ベンリー強制反映)
// 更新: 2026-05-13 すい(凪設計) v6.0 [C-063 Phase1] - #T-XXX 経路追加（既存温存）
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
  if (event?.source?.type !== 'group') return false;

  const groupId = event.source.groupId;
  const rawText = String(text || '');

  // [C-063 Phase1] 先頭行から #T-XXX を抽出（あれば本文から除外）
  const lines = rawText.split(/\r?\n/);
  let storeId = null;
  let bodyLines = lines;

  if (lines.length > 0) {
    const firstLine = lines[0].trim();
    const m = firstLine.match(/^#?\s*T[\-]?(\d{3,})\s*$/i);
    if (m) {
      storeId = `T-${m[1]}`;
      bodyLines = lines.slice(1);
    }
  }

  const body = bodyLines.join('\n').trim();

  // [C-063 Phase1] 既存パス（SHIFT_GROUP_ID）と新パス（T-XXX指定）の共存
  // storeId あり → どのグループからでも T-XXX で店舗特定
  // storeId なし → 既存 SHIFT_GROUP_ID 一致のみ通過（CREA運用温存）
  if (!storeId && groupId !== SHIFT_GROUP_ID) {
    return false;
  }

  const parts = body.split(/[\s　]+/);
  // すいさん v5.1 緩和: parts.length >= 2 で反応 (旧 >= 3 から緩和)
  if (parts.length < 2 && body !== '確認' && body !== '出勤確認') {
    return false;
  }

  setImmediate(async () => {
    try {
      // [C-063 Phase1] T-XXX 指定の場合は権限チェック（許可LINE userId検証）
      if (storeId) {
        try {
          const { getFolderByStoreId } = require('./line_nippo_input');
          const userId = event.source?.userId;
          await getFolderByStoreId(storeId, userId);
          console.log(`[Shift] [C-063] ${storeId} 権限OK userId=${userId}`);
        } catch (err) {
          console.error('[Shift] [C-063] 権限エラー:', err.message);
          await client.pushMessage({
            to: groupId,
            messages: [{ type: 'text', text: `❌ ${err.message}` }],
          }).catch(() => {});
          return;
        }
        // [C-063 Phase2 で実装] storeId → 店舗別シフトSS の選択ロジック
        // 現状（Phase1）は shift_reflect.reflectShiftMessage が CREA+ふわもこ前提のため、
        // T-001 以外の店舗指定でも CREA+ふわもこのシフトSSにマッチを試みる。
        // T-001 以外で運用するなら Phase2 完了まで待つこと。
      }

      const shiftReflect = require('./shift_reflect');
      const result = await shiftReflect.reflectShiftMessage(body);
      if (result.type === 'ignore') {
        console.log('[Shift] 非シフトメッセージ→スルー');
      } else if (result.type === 'success') {
        console.log(`[Shift] 結果: success ${result.writtenCount} 件 (${result.staffName} / ${result.store})${storeId ? ` storeId=${storeId}` : ''}`);
        // C-040 連携 (v5.4): SS 書込成功スタッフを selected_staff として workflow_dispatch に渡し、
        // 既存値と同じでもベンリーに強制反映させる
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
          to: groupId,
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
