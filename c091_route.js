// /app/chatbot/c091_route.js - C-091 シフト自動反映SaaS 新台帳ルート
// 2026-07-02 かず(凪代行) 新規作成
//
// ★役割: 設定ファイルに登録された「テスト用LINEグループ」からのシフト投稿だけを
//   新台帳(C-091)へ書き込む。既存の shift_route.js(C-036/すい管轄) には一切触れない。
//
// ★安全設計:
//   - /app/c091/route_config.json の groups に groupId が無ければ完全に不活性(return false)
//     → 設定が空の間は既存運用への影響ゼロ。
//   - シフト形式(1行目=キャスト名/2行目以降に日付行)でなければ不活性 → 雑談は素通し。
//   - 書込本体は /app/c091/ledger_writer_2026_06_25.js(実績あるCLI)へ子プロセス委譲。
//   - 名簿は roster_from_ledger.js で「台帳A列の現在順」から毎回生成
//     → 既存行の並びは絶対に変わらない(かず指示2026-07-02の不変ルール遵守)。
//
// 設定例: {"groups": {"Cxxxxxxxx...": "T-001"}}
const fs = require('fs');
const { execFile } = require('child_process');

const CONFIG_PATH = '/app/c091/route_config.json';
const PARSER_PATH = '/app/c091/shift_parser_v2_2026_06_25.js';
const ROSTER_GEN = '/app/c091/roster_from_ledger.js';
const WRITER_PATH = '/app/c091/ledger_writer_2026_06_25.js';
const C091_ENV = { ...process.env, NODE_PATH: '/app/c091/node_modules' };

function loadConfig() {
  try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')); } catch (_) { return null; }
}

// JSTの今日 (サーバーはUTC)
function jstTodayStr() {
  const d = new Date(Date.now() + 9 * 3600 * 1000);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
}

function run(args) {
  return new Promise((resolve, reject) => {
    execFile('node', args, { env: C091_ENV, timeout: 90000, maxBuffer: 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) reject(new Error((stderr || stdout || err.message).trim()));
      else resolve((stdout || '').trim());
    });
  });
}

/**
 * @returns {Promise<boolean>} true=このルートで処理した / false=素通し(既存処理へ)
 */
async function handle(event, text, client) {
  if (event?.source?.type !== 'group') return false;
  const cfg = loadConfig();
  const storeId = cfg && cfg.groups && cfg.groups[event.source.groupId];
  if (!storeId) return false;

  // シフト形式判定: 1行目=キャスト名(日付でない) + 2行目以降に日付行が1つ以上
  let P;
  try { P = require(PARSER_PATH); } catch (e) { console.error('[c091_route] parser読込エラー:', e.message); return false; }
  const lines = String(text || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  if (lines.length < 2) return false;
  if (P.parseLine(lines[0])) return false;                       // 1行目が日付 → 想定形式でない
  if (!lines.slice(1).some(l => P.parseLine(l))) return false;   // 日付行なし → 雑談として素通し

  const replyToken = event.replyToken;
  const groupId = event.source.groupId;
  console.log(`[c091_route] シフト投稿受信 store=${storeId} group=${groupId} 1行目="${lines[0]}"`);

  setImmediate(async () => {
    const sendBack = async (msg) => {
      try { await client.replyMessage({ replyToken, messages: [{ type: 'text', text: msg }] }); }
      catch (_) { await client.pushMessage({ to: groupId, messages: [{ type: 'text', text: msg }] }).catch(() => {}); }
    };
    try {
      const today = jstTodayStr();
      const rosterPath = `/tmp/c091_roster_${storeId}.json`;
      await run([ROSTER_GEN, '--out', rosterPath]);              // 名簿=台帳A列の現在順(並び不変)

      const msgPath = `/tmp/c091_msg_${Date.now()}_${Math.floor(Math.random() * 1e6)}.txt`;
      fs.writeFileSync(msgPath, lines.join('\n'));
      const out = await run([WRITER_PATH, '--file', msgPath, '--today', today, '--roster', rosterPath, '--commit']);
      fs.unlink(msgPath, () => {});

      // 返信サマリ: writer出力から要点行を抽出
      const summary = out.split('\n')
        .filter(l => /^(キャスト:|解析:|  \[)/.test(l))
        .join('\n') || out.slice(-300);
      console.log(`[c091_route] 書込完了 store=${storeId}\n${summary}`);
      await sendBack(`✅ シフト表(新台帳)に反映しました\n${summary}`);
    } catch (e) {
      // 名簿外・候補複数・タブ未作成などは writer が ❌ 行で説明を出す → そのまま返す
      const errLine = String(e.message || '').split('\n').filter(l => l.includes('❌')).join('\n') || 'エラーが発生しました（管理者にご連絡ください）';
      console.error('[c091_route] エラー:', e.message);
      await sendBack(`⚠️ 台帳に反映できませんでした\n${errLine}`);
    }
  });
  return true;
}

module.exports = { handle };
