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
const ESTAMA_KANBAI = '/app/c091/venrey/venrey-automation/estama_kanbai_c091.py'; // [2026-07-04 かず] 当欠/無欠/店欠→エステ魂完売
const ESTAMA_CREDS = '/app/c091/estama_creds_c091.js';                            // [2026-07-04 かず] エステ魂認証取得
const PYBIN = '/app/c091/venv/bin/python';
const PYCWD = '/app/c091/venrey/venrey-automation';
const C091_ENV = { ...process.env, NODE_PATH: '/app/c091/node_modules', PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8', HEADLESS: 'true' };

function loadConfig() {
  try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')); } catch (_) { return null; }
}

// [2026-07-02 かず] 直列処理キュー: 複数キャストが同時にシフトを送っても1件ずつ処理する。
// 並行処理すると ①台帳writerの全体書き戻しで後勝ち消失(先の子のシフトが消える)
// ②ヘッドレスChrome多重起動でVPS(2コア/4GB)が不安定 になるため必須。
let _queueTail = Promise.resolve();
let _queueLen = 0;
function enqueue(task) {
  _queueLen++;
  const pos = _queueLen;
  if (pos > 1) console.log(`[c091_route] キュー待ち ${pos - 1}件あり → 順番に処理します`);
  const next = _queueTail.then(task, task).finally(() => { _queueLen--; });
  _queueTail = next.catch(() => {});
  return next;
}

// [2026-07-02 かず] スプール永続化: 受信したシフト投稿をディスクに保存し、処理完了で削除する。
// pm2 restart・クラッシュ・再起動でキュー処理が中断されても、次の受信時に自動復旧する
// (2026-07-02 りなの投稿が再起動で消失した実害の再発防止)。
const SPOOL_DIR = '/app/c091/_spool';
function spoolWrite(groupId, text) {
  try {
    fs.mkdirSync(SPOOL_DIR, { recursive: true });
    const p = `${SPOOL_DIR}/${Date.now()}_${Math.floor(Math.random() * 1e6)}.json`;
    fs.writeFileSync(p, JSON.stringify({ groupId, text, ts: new Date().toISOString() }));
    return p;
  } catch (e) { console.error('[c091_route] spool書込失敗:', e.message); return null; }
}

let _spoolRecovered = false;
function recoverSpool(client) {
  if (_spoolRecovered) return;
  _spoolRecovered = true;
  let files = [];
  try { files = fs.readdirSync(SPOOL_DIR).filter(f => f.endsWith('.json')).sort(); } catch (_) { return; }
  for (const f of files) {
    const p = `${SPOOL_DIR}/${f}`;
    try {
      const j = JSON.parse(fs.readFileSync(p, 'utf8'));
      const lines = String(j.text || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      const cfgAll = loadConfig();
      const gv = cfgAll && cfgAll.groups && cfgAll.groups[j.groupId];
      if (!gv || lines.length < 2) { fs.unlink(p, () => {}); continue; }
      const st = Array.isArray(gv) ? gv : [gv];
      console.log(`[c091_route] 再起動で中断された投稿を復旧処理: 1行目="${lines[0]}"`);
      enqueue(() => processShiftPost(client, st, lines, j.groupId, null, p));
    } catch (e) { console.error('[c091_route] spool復旧失敗:', f, e.message); }
  }
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
    execFile(cmd, args, { env: opts.env || C091_ENV, timeout: opts.timeout || 90000, maxBuffer: 1024 * 1024, cwd: opts.cwd }, (err, stdout, stderr) => {
      if (err) reject(new Error((stderr || stdout || err.message).trim()));
      else resolve((stdout || '').trim());
    });
  });
}

// [2026-07-19 かず] Sheets APIの一時的なタイムアウト/レート制限に備え、設定・名簿の取得はリトライする。
//   これが無いと1回の一時失敗で resolveStore がその店舗を黙って候補から外し、
//   実在キャスト(例:小林もえ)を「名簿にありません」と誤返信していた(T-001名簿取得タイムアウトで実害)。
async function runRetry(cmd, args, tries = 3) {
  let lastErr;
  for (let i = 0; i < tries; i++) {
    try { return await run(cmd, args); }
    catch (e) { lastErr = e; if (i < tries - 1) await new Promise(r => setTimeout(r, 1500 * (i + 1))); }
  }
  throw lastErr;
}

// 対象店舗群の名簿から店舗+キャストを特定(完全一致優先→部分一致・一意でなければ聞き返し)
async function resolveStore(storeTags, firstLine) {
  const t = norm(firstLine);
  const exact = [], partial = [];
  const failedTags = [];   // 名簿(または設定)を取得できなかった店舗。名前の有無を判定できないので「無い」と断定しない
  for (const tag of storeTags) {
    let cfg;
    try { cfg = JSON.parse(await runRetry('node', [SETTINGS_TOOL, 'config', tag])); }
    catch (e) { console.error(`[c091_route] ${tag} 設定取得失敗:`, e.message); failedTags.push(tag); continue; }
    if (!cfg.ledgerSsId) continue;
    const rosterPath = `/tmp/c091_roster_${tag}.json`;
    try { await runRetry('node', [ROSTER_GEN, '--out', rosterPath, '--ss', cfg.ledgerSsId]); }
    catch (e) { console.error(`[c091_route] ${tag} 名簿取得失敗:`, e.message); failedTags.push(tag); continue; }
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
  // ヒット無し。ただし名簿を取得できなかった店舗がある場合は「名前が無い」と誤断せず、一時失敗として通知する。
  if (failedTags.length) return { rosterError: failedTags };
  return {};
}

// [2026-07-06 かず指示] 「前欠」を日付なしで送った時の聞き返し検知。
//   仕様: 前欠は日付必須。日付なし(例「立花 前欠」「立花⏎前欠」)は処理せず⚠️で日付を促す。
//   誤反応防止のため厳密判定: 名前行末尾がちょうど前欠 or 2行目以降がすべて「前欠」だけ、の時のみ候補名を返す。
//   雑談(例「立花さん前欠かも」)は末尾が前欠でないので非該当。日付付き前欠は通常処理へ流す(nullを返す)。
function bareZenketsuName(lines, P) {
  const isBareZK = s => s.replace(/\s+/g, '') === '前欠';
  if (lines.length === 1) {
    const m = lines[0].match(/^(.+?)\s*前欠\s*$/);
    return (m && m[1].trim() && !/前欠/.test(m[1])) ? m[1].trim() : null;
  }
  if (lines.slice(1).some(l => P.parseLine(l))) return null;      // 日付付き行あり→通常処理へ
  const rest = lines.slice(1);
  if (!(rest.length && rest.every(isBareZK))) return null;        // 2行目以降が全部「前欠」だけの時のみ
  const m0 = lines[0].match(/^(.+?)\s*前欠\s*$/);
  return (m0 ? m0[1].trim() : lines[0].trim()) || null;
}

// 前欠(日付なし)を、本物のキャスト名に完全一致した時だけ⚠️聞き返す。返り値: 聞き返したらtrue。
async function askZenketsuDate(candName, storeTags, event, client) {
  let r;
  try { r = await resolveStore(storeTags, candName); } catch (e) { return false; }
  const t = norm(candName);
  const exactHit = (r.hit && norm(r.hit.name) === t) ? r.hit
    : (r.ambiguous && r.ambiguous.find(h => norm(h.name) === t)) || null;
  if (!exactHit) return false;                                    // 名簿に完全一致しない→誤反応防止で素通し
  const msg = `⚠️「前欠」は日付を付けて送ってください\n例：\n${exactHit.name}\n7/6 前欠\n（前欠は日にちが必要です）`;
  try { await client.replyMessage({ replyToken: event.replyToken, messages: [{ type: 'text', text: msg }] }); }
  catch (_) { await client.pushMessage({ to: event.source.groupId, messages: [{ type: 'text', text: msg }] }).catch(() => {}); }
  console.log(`[c091_route] 前欠(日付なし)聞き返し: ${exactHit.name}`);
  return true;
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
  recoverSpool(client);   // 対象グループの着信を契機に、再起動で中断された投稿を復旧(初回のみ)

  // シフト形式判定: 1行目=キャスト名(日付でない) + 2行目以降に日付行が1つ以上
  let P;
  try { P = require(PARSER_PATH); } catch (e) { console.error('[c091_route] parser読込エラー:', e.message); return false; }
  const lines = String(text || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  if (lines.length < 2) {
    // 1行「立花 前欠」形式: 前欠(日付なし)なら聞き返し、それ以外は素通し
    const cand = bareZenketsuName(lines, P);
    if (cand && await askZenketsuDate(cand, storeTags, event, client)) return true;
    return false;
  }
  if (P.parseLine(lines[0])) return false;                       // 1行目が日付 → 想定形式でない
  if (!lines.slice(1).some(l => P.parseLine(l))) {               // 日付行なし
    // 前欠(日付なし・例「立花⏎前欠」)なら日付を促す聞き返し。雑談は素通し。
    const cand = bareZenketsuName(lines, P);
    if (cand && await askZenketsuDate(cand, storeTags, event, client)) return true;
    return false;                                                // それ以外は雑談として素通し
  }

  const replyToken = event.replyToken;
  const groupId = event.source.groupId;
  console.log(`[c091_route] シフト投稿受信 stores=${storeTags.join('/')} group=${groupId} 1行目="${lines[0]}"`);

  const spoolPath = spoolWrite(groupId, text);     // 処理完了までディスクに保持(消失防止)
  enqueue(() => processShiftPost(client, storeTags, lines, groupId, replyToken, spoolPath));
  return true;
}

// シフト投稿1件の処理本体(キューから1件ずつ実行される)
async function processShiftPost(client, storeTags, lines, groupId, replyToken, spoolPath) {
    let P;
    try { P = require(PARSER_PATH); } catch (e) { console.error('[c091_route] parser読込エラー:', e.message); return; }
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
      if (r.rosterError) {
        // 名簿を取得できなかった店舗があり、名前の有無を判定できない状態。名前ミスと誤解させないよう明示する。
        console.error(`[c091_route] 名簿一時取得失敗のため保留: 失敗店舗=${r.rosterError.join('/')} name="${lines[0]}"`);
        await sendBack(`⚠️システムが一時的に名簿を取得できませんでした(対象: ${r.rosterError.join('/')})。\n名前の間違いではありません。30秒ほどおいて、もう一度そのまま送ってください🙇`);
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
      // [2026-07-02 かず指示] 成功時のLINE返信は送らない(✅台帳反映の返信を廃止)。⚠️系(聞き返し/失敗)のみ残す

      // ── 2) ベンリー即時反映(その子だけ) ──
      // [2026-07-02 かず指示] ベンリーに入れるのは「今日を含めて1週間(今日〜+6日)」だけ。
      // それより先・過去の日付は台帳のみ(未来分は各当日の深夜cron nightly_reflect が反映)。
      const winStart = today.getTime();
      const winEnd = today.getTime() + 6 * 86400000;
      const entries = [];
      const deferred = [];
      const ledgerOnly = [];
      for (const l of lines.slice(1)) {
        const e = P.parseLine(l);
        if (!e || !e.venrey || e.venrey.action === 'skip') continue;
        const ymd = P.determineYearMonth(today, e);
        // [2026-07-04 かず指示] 当欠・無欠は台帳のみ。ベンリーは触らない(更新も配信もしない)
        if (e.venrey.action === 'ledger_only') { ledgerOnly.push({ md: `${ymd.month}/${ymd.day}`, category: e.category }); continue; }
        const dt = new Date(ymd.year, ymd.month - 1, ymd.day).getTime();
        if (dt < winStart || dt > winEnd) { deferred.push(`${ymd.month}/${ymd.day}`); continue; }
        entries.push({
          date: `${ymd.year}-${String(ymd.month).padStart(2, '0')}-${String(ymd.day).padStart(2, '0')}`,
          action: e.venrey.action, start: e.venrey.start || null, end: e.venrey.end || null, raw: e.ss,
        });
      }
      if (deferred.length) console.log(`[c091_route] 1週間より先/過去のためベンリー即時反映せず(台帳のみ・当日深夜cronに委譲): ${deferred.join(', ')}`);
      if (ledgerOnly.length) {
        console.log(`[c091_route] 当欠/無欠/店欠のため台帳のみ・ベンリー触らず: ${ledgerOnly.map(o => `${o.md}(${o.category})`).join(', ')}`);
        // [2026-07-04 かず指示] エステ魂だけは該当日を【完売】(全枠×)にして予約を止める(シフト表示は残す)。
        //   ベンリーは引き続き触らない。エステ魂側が失敗してもLINE処理・台帳記入は止めない。
        //   店側が既にお休み処理済み等で完売ボタンが押せない日はスクリプト側が⏭スキップ(正常終了)する。
        for (const o of ledgerOnly) {
          try {
            const cred = (await run('node', [ESTAMA_CREDS, storeId])).trim().split('\t');
            if (cred.length < 2 || !cred[0] || !cred[1]) throw new Error('エステ魂の認証情報を取得できません');
            const kOut = await run(PYBIN, [ESTAMA_KANBAI, storeId, resolved, o.md, '--commit'],
              { cwd: PYCWD, timeout: 300000, env: { ...C091_ENV, ESTAMA_USER: cred[0], ESTAMA_PASS: cred[1] } });
            console.log(`[c091_route] エステ魂完売 ${storeId} ${resolved} ${o.md}: ${kOut.split('\n').pop()}`);
          } catch (err) {
            console.error(`[c091_route] エステ魂完売エラー ${storeId} ${resolved} ${o.md}:`, String(err.message).slice(0, 300));
            await push(`⚠️ ${resolved} ${o.md}(${o.category}): エステ魂の完売反映に失敗しました（台帳記入は完了しています。手動で完売にしてください）`).catch(() => {});
          }
        }
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

      // ── 4) 結果通知 ── [2026-07-02 かず指示] 成功時は返信しない。実失敗(表示範囲外以外)がある時だけ⚠️通知
      const realNg = (res.skipped || []).filter(s => s.reason !== 'ベンリーの表示範囲外');
      if (realNg.length) {
        const ngLines = realNg.map(s => `  ${s.date.slice(5).replace('-', '/')} ⚠️${s.reason}`);
        await push(`⚠️ ${resolved}: ベンリーに反映できなかった日があります\n${ngLines.join('\n')}`);
      }
    } catch (e) {
      const errLine = String(e.message || '').split('\n').filter(l => l.includes('❌')).join('\n') || 'エラーが発生しました（管理者にご連絡ください）';
      console.error('[c091_route] エラー:', e.message);
      await sendBack(`⚠️ 反映できませんでした\n${errLine}`);
    } finally {
      if (spoolPath) fs.unlink(spoolPath, () => {});   // 処理が終わった(成功/失敗問わず応答済み)のでスプール削除
    }
}

module.exports = { handle };
