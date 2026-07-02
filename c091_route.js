// /app/chatbot/c091_route.js - C-091 シフト自動反映SaaS 新台帳ルート
// 2026-07-02 かず(凪代行) 新規作成
// 2026-07-02 v2: LINE投稿→台帳書込に加え、その子だけベンリー即時反映+配信+水色マークを追加(かず指示)
// 2026-07-02 v3: 多店舗対応(設定シートから台帳SS_IDを解決)
// 2026-07-02 v4: 1グループ複数店舗対応(かず指示: CREA=苗字/ふわもこ=下の名前で同一グループ受付)
//   - route_config.json の groups 値は "T-001" または ["T-001","T-002"](配列=両店受付)
//   - 店舗特定: 全対象店舗の名簿(台帳A列)で ①完全一致優先 ②次に部分一致。
//     一意に決まらなければ書き込まずに聞き返す(例: るな→ふわもこ「るな」に確定、CREAは「星野」で送る)
//
// ★安全設計:
//   - groups に groupId が無ければ完全に不活性(return false)
//   - シフト形式(1行目=キャスト名/2行目以降に日付行)でなければ不活性 → 雑談は素通し
//   - 台帳書込は ledger_writer、ベンリー反映は apply_cast_c091.py(ID照合つき)へ子プロセス委譲
//   - 名簿は roster_from_ledger.js で「台帳A列の現在順」から毎回生成(既存行の並び不変)
const fs = require('fs');
const { execFile } = require('child_process');

const CONFIG_PATH = '/app/c091/route_config.json';
const PARSER_PATH = '/app/c091/shift_parser_v2_2026_06_25.js';
const ROSTER_GEN = '/app/c091/roster_from_ledger.js';
const WRITER_PATH = '/app/c091/ledger_writer_2026_06_25.js';
const MARK_PATH = '/app/c091/mark_ledger_status.js';
const SETTINGS_TOOL = '/app/c091/c091_settings_tool.js';
const APPLY_CAST = '/app/c091/venrey/venrey-automation/apply_cast_c091.py';
const PYBIN = '/app/c091/venv/bin/python';
const PYCWD = '/app/c091/venrey/venrey-automation';
const C091_ENV = { ...process.env, NODE_PATH: '/app/c091/node_modules', PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8', HEADLESS: 'true' };

function loadConfig() {
  try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')); } catch (_) { return null; }
}

// JSTの今日 (サーバーはUTC)
function jstToday() {
  const d = new Date(Date.now() + 9 * 3600 * 1000);
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}
const isoOf = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const norm = s => String(s || '').replace(/[\s　]/g, '');

function run(cmd, args, opts = {}) {
  return new Promise((resolve, reject) => {
    execFile(cmd, args, { env: C091_ENV, timeout: opts.timeout || 90000, maxBuffer: 1024 * 1024, cwd: opts.cwd }, (err, stdout, stderr) => {
      if (err) reject(new Error((stderr || stdout || err.message).trim()));
      else resolve((stdout || '').trim());
    });
  });
}

// 対象店舗群の名簿から店舗+キャストを特定(完全一致優先→部分一致・一意でなければ聞き返し)
async function resolveStore(storeTags, firstLine) {
  const t = norm(firstLine);
  const exact = [], partial = [];
  for (const tag of storeTags) {
    let cfg;
    try { cfg = JSON.parse(await run('node', [SETTINGS_TOOL, 'config', tag])); }
    catch (e) { console.error(`[c091_route] ${tag} 設定取得失敗:`, e.message); continue; }
    if (!cfg.ledgerSsId) continue;
    const rosterPath = `/tmp/c091_roster_${tag}.json`;
    try { await run('node', [ROSTER_GEN, '--out', rosterPath, '--ss', cfg.ledgerSsId]); }
    catch (e) { console.error(`[c091_route] ${tag} 名簿取得失敗:`, e.message); continue; }
    const roster = (JSON.parse(fs.readFileSync(rosterPath, 'utf8')).roster || []);
    for (const r of roster) {
      const n = norm(r.name);
      if (!n) continue;
      if (n === t) exact.push({ tag, name: r.name, cfg, rosterPath });
      else if (n.includes(t) || t.includes(n)) partial.push({ tag, name: r.name, cfg, rosterPath });
    }
  }
  if (exact.length === 1) return { hit: exact[0] };
  if (exact.length > 1) return { ambiguous: exact };
  if (partial.length === 1) return { hit: partial[0] };
  if (partial.length > 1) return { ambiguous: partial };
  return {};
}

/**
 * @returns {Promise<boolean>} true=このルートで処理した / false=素通し(既存処理へ)
 */
async function handle(event, text, client) {
  if (event?.source?.type !== 'group') return false;
  const cfgAll = loadConfig();
  const groupVal = cfgAll && cfgAll.groups && cfgAll.groups[event.source.groupId];
  if (!groupVal) return false;
  const storeTags = Array.isArray(groupVal) ? groupVal : [groupVal];

  // シフト形式判定: 1行目=キャスト名(日付でない) + 2行目以降に日付行が1つ以上
  let P;
  try { P = require(PARSER_PATH); } catch (e) { console.error('[c091_route] parser読込エラー:', e.message); return false; }
  const lines = String(text || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  if (lines.length < 2) return false;
  if (P.parseLine(lines[0])) return false;                       // 1行目が日付 → 想定形式でない
  if (!lines.slice(1).some(l => P.parseLine(l))) return false;   // 日付行なし → 雑談として素通し

  const replyToken = event.replyToken;
  const groupId = event.source.groupId;
  console.log(`[c091_route] シフト投稿受信 stores=${storeTags.join('/')} group=${groupId} 1行目="${lines[0]}"`);

  setImmediate(async () => {
    const push = async (msg) => {
      await client.pushMessage({ to: groupId, messages: [{ type: 'text', text: msg }] }).catch((e) => console.error('[c091_route] push失敗:', e.message));
    };
    const sendBack = async (msg) => {
      try { await client.replyMessage({ replyToken, messages: [{ type: 'text', text: msg }] }); }
      catch (_) { await push(msg); }
    };
    try {
      // ── 0) 店舗+キャストを特定 ──
      const r = await resolveStore(storeTags, lines[0]);
      if (r.ambiguous) {
        const cand = r.ambiguous.map(h => `${h.cfg.name}「${h.name}」`).join(' / ');
        await sendBack(`⚠️「${lines[0]}」は複数の候補があります: ${cand}\nフルネーム(または店舗が分かる名前)で送り直してください`);
        return;
      }
      if (!r.hit) {
        await sendBack(`⚠️「${lines[0]}」がどの店舗の名簿にも見つかりません(対象: ${storeTags.join('/')})\n名前を確認して送り直してください`);
        return;
      }
      const { tag: storeId, name: resolved, cfg: storeCfg, rosterPath } = r.hit;
      const ledgerSs = storeCfg.ledgerSsId;
      console.log(`[c091_route] 店舗特定: ${storeId}(${storeCfg.name}) キャスト=${resolved}`);

      // ── 1) 台帳書込 (1行目を確定名に差し替えて渡す) ──
      const today = jstToday();
      const msgPath = `/tmp/c091_msg_${Date.now()}_${Math.floor(Math.random() * 1e6)}.txt`;
      fs.writeFileSync(msgPath, [resolved, ...lines.slice(1)].join('\n'));
      const out = await run('node', [WRITER_PATH, '--file', msgPath, '--today', isoOf(today), '--roster', rosterPath, '--ss', ledgerSs, '--commit']);
      fs.unlink(msgPath, () => {});

      const summary = out.split('\n')
        .filter(l => /^(キャスト:|解析:|  \[)/.test(l))
        .join('\n') || out.slice(-300);
      console.log(`[c091_route] 書込完了 store=${storeId}\n${summary}`);
      await sendBack(`✅ ${storeCfg.name} のシフト表に反映しました\nキャスト: ${resolved}\n${summary.split('\n').filter(l => !l.startsWith('キャスト:')).join('\n')}\n⏳続けてベンリーを更新中…(1〜2分)`);

      // ── 2) ベンリー即時反映(その子だけ) ──
      const entries = [];
      for (const l of lines.slice(1)) {
        const e = P.parseLine(l);
        if (!e || !e.venrey || e.venrey.action === 'skip') continue;
        const ymd = P.determineYearMonth(today, e);
        entries.push({
          date: `${ymd.year}-${String(ymd.month).padStart(2, '0')}-${String(ymd.day).padStart(2, '0')}`,
          action: e.venrey.action, start: e.venrey.start || null, end: e.venrey.end || null, raw: e.ss,
        });
      }
      if (!entries.length) return;

      const entPath = `/tmp/c091_entries_${Date.now()}.json`;
      fs.writeFileSync(entPath, JSON.stringify(entries));
      let res = null;
      try {
        const pyOut = await run(PYBIN, [APPLY_CAST, storeId, resolved, entPath], { cwd: PYCWD, timeout: 300000 });
        const m = pyOut.match(/JSON_RESULT: (\{.*\})/);
        res = m ? JSON.parse(m[1]) : null;
        console.log(`[c091_route] ベンリー反映結果 ${storeId} ${resolved}:`, m ? m[1] : pyOut.slice(-200));
      } finally {
        fs.unlink(entPath, () => {});
      }
      if (!res) { await push(`⚠️ ${resolved}: ベンリー反映の結果を確認できませんでした（管理者にご連絡ください）`); return; }

      // ── 3) 反映できた日付セルを水色マーク ──
      for (const a of res.applied) {
        const [y, mo, d] = a.date.split('-').map(Number);
        await run('node', [MARK_PATH, '--date', `${mo}/${d}`, '--mode', 'bg', '--names', resolved, '--today', isoOf(today), '--ss', ledgerSs])
          .catch(e => console.error('[c091_route] 水色マーク失敗:', e.message));
      }

      // ── 4) 結果をLINEへ ──
      const okLines = res.applied.map(a => `  ${a.date.slice(5).replace('-', '/')} ${a.label}`);
      const ngLines = (res.skipped || []).map(s => `  ${s.date.slice(5).replace('-', '/')} ⚠️${s.reason}`);
      let msg = `🚀 ベンリー更新完了: ${storeCfg.name} ${resolved}`;
      if (okLines.length) msg += `\n${okLines.join('\n')}\n📡 サイトへ配信済み(エステ魂の反映は毎時自動確認→赤文字)`;
      if (ngLines.length) msg += `\n${ngLines.join('\n')}`;
      if (!okLines.length) msg = `⚠️ ${resolved}: ベンリーに反映できませんでした\n${ngLines.join('\n')}`;
      await push(msg);
    } catch (e) {
      const errLine = String(e.message || '').split('\n').filter(l => l.includes('❌')).join('\n') || 'エラーが発生しました（管理者にご連絡ください）';
      console.error('[c091_route] エラー:', e.message);
      await sendBack(`⚠️ 反映できませんでした\n${errLine}`);
    }
  });
  return true;
}

module.exports = { handle };
