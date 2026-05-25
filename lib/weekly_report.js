// C-060 拡張：VPS週次レポート
// KIRAKU の crontab から毎週月曜 00:00 UTC (JST 09:00) に実行される。
// 自身（KIRAKU）と エスタマ(159.69.40.60) の過去7日 sar データを集計して
// 監視LINEグループに1通でまとめて通知する。
//
// env: LINE_CHANNEL_ACCESS_TOKEN, LINE_MONITOR_GROUP_ID
// 依存: /root/.ssh/id_ed25519_estama (KIRAKU→エスタマ片方向SSH鍵)

'use strict';

const { execSync } = require('child_process');
const https = require('https');

const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const LINE_GROUP = process.env.LINE_MONITOR_GROUP_ID;
const ESTAMA_HOST = '159.69.40.60';
const ESTAMA_SSH_KEY = '/root/.ssh/id_ed25519_estama';

function sh(cmd) {
  return execSync(cmd, { encoding: 'utf8', timeout: 30000 }).trim();
}

function shRemote(cmd) {
  const safe = cmd.replace(/'/g, `'\\''`);
  return sh(`ssh -o StrictHostKeyChecking=no -o ConnectTimeout=15 -i ${ESTAMA_SSH_KEY} root@${ESTAMA_HOST} '${safe}'`);
}

// 過去7日分のsysstatファイル名 (sa01..sa31) を返す
function last7SarFiles() {
  const files = [];
  for (let i = 1; i <= 7; i++) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const day = String(d.getDate()).padStart(2, '0');
    files.push(`/var/log/sysstat/sa${day}`);
  }
  return files;
}

// sadf -d でCSV出力された数値列をパースし、平均とピークを返す
function parseValues(out) {
  const values = out
    .split('\n')
    .filter((l) => l && !l.startsWith('#'))
    .map((l) => parseFloat(l))
    .filter((v) => isFinite(v));
  if (!values.length) return { avg: null, peak: null, n: 0 };
  const avg = values.reduce((s, v) => s + v, 0) / values.length;
  return { avg, peak: Math.max(...values), n: values.length };
}

// CPU使用率（user+nice+system+iowait+steal = 100-idle）
function getCpuStats(runner, file) {
  try {
    const out = runner(
      `test -f ${file} && sadf -d ${file} -- -u 2>/dev/null | awk -F';' '!/^#/ && NF>=10 {printf "%.2f\\n", 100-$10}' || true`
    );
    return parseValues(out);
  } catch (_) {
    return { avg: null, peak: null, n: 0 };
  }
}

// ldavg-1（直近1分）
function getLoadStats(runner, file) {
  try {
    const out = runner(
      `test -f ${file} && sadf -d ${file} -- -q 2>/dev/null | awk -F';' '!/^#/ && NF>=6 {print $6}' || true`
    );
    return parseValues(out);
  } catch (_) {
    return { avg: null, peak: null, n: 0 };
  }
}

// %memused
function getMemStats(runner, file) {
  try {
    const out = runner(
      `test -f ${file} && sadf -d ${file} -- -r 2>/dev/null | awk -F';' '!/^#/ && NF>=7 {print $7}' || true`
    );
    return parseValues(out);
  } catch (_) {
    return { avg: null, peak: null, n: 0 };
  }
}

function uptimeDays(runner) {
  try {
    const out = runner(`awk '{print int($1/86400)}' /proc/uptime`);
    const v = parseInt(out, 10);
    return isFinite(v) ? v : null;
  } catch (_) {
    return null;
  }
}

function diskPercent(runner) {
  try {
    const out = runner(`df -h / | awk 'NR==2 {gsub("%","",$5); print $5}'`);
    const v = parseFloat(out);
    return isFinite(v) ? v : null;
  } catch (_) {
    return null;
  }
}

// 7日分のファイルを結合して全データから平均/ピーク算出
function collect(label, runner) {
  const files = last7SarFiles();
  // 全7日分の生の値をかき集める
  const cpuAll = [];
  const loadAll = [];
  const memAll = [];
  let daysWithData = 0;

  for (const f of files) {
    const cpu = getCpuStats(runner, f);
    const load = getLoadStats(runner, f);
    const mem = getMemStats(runner, f);
    if (cpu.n > 0 || load.n > 0 || mem.n > 0) daysWithData++;

    // 個別ファイルのavgとpeakをそのまま使うのではなく、全期間で計算したい
    // そのため、各ファイルからavg値を保持（後で再平均）し、peakは最大値を取る
    if (cpu.avg != null) cpuAll.push({ avg: cpu.avg, peak: cpu.peak, n: cpu.n });
    if (load.avg != null) loadAll.push({ avg: load.avg, peak: load.peak, n: load.n });
    if (mem.avg != null) memAll.push({ avg: mem.avg, peak: mem.peak, n: mem.n });
  }

  // 加重平均（観測数で重み付け）とピークの最大値
  const weighted = (arr) => {
    if (!arr.length) return { avg: null, peak: null };
    const totalN = arr.reduce((s, x) => s + x.n, 0);
    if (!totalN) return { avg: null, peak: null };
    const avg = arr.reduce((s, x) => s + x.avg * x.n, 0) / totalN;
    const peak = Math.max(...arr.map((x) => x.peak));
    return { avg, peak };
  };

  const cpu = weighted(cpuAll);
  const load = weighted(loadAll);
  const mem = weighted(memAll);

  return {
    label,
    days: daysWithData,
    cpu_avg: cpu.avg,
    cpu_peak: cpu.peak,
    load_avg: load.avg,
    load_peak: load.peak,
    mem_avg: mem.avg,
    uptime_days: uptimeDays(runner),
    disk_percent: diskPercent(runner),
  };
}

function fmt(v, digits = 1) {
  return v == null ? '?' : Number(v).toFixed(digits);
}

function buildMessage(kiraku, estama) {
  const today = new Date();
  const start = new Date();
  start.setDate(today.getDate() - 7);
  const fmtDate = (d) => `${d.getMonth() + 1}/${d.getDate()}`;
  const period = `${fmtDate(start)}-${fmtDate(today)}`;

  const block = (s) =>
    [
      `✅ 稼働日数: ${s.uptime_days != null ? s.uptime_days + '日連続' : '?'}`,
      `📈 CPU平均: ${fmt(s.cpu_avg)}% / ピーク ${fmt(s.cpu_peak)}%`,
      `⚡ Load平均: ${fmt(s.load_avg, 2)} / ピーク ${fmt(s.load_peak, 2)}`,
      `💾 メモリ平均: ${fmt(s.mem_avg)}%`,
      `💿 ディスク: ${fmt(s.disk_percent)}%`,
      `(集計対象: 過去${s.days}日分のsar)`,
    ].join('\n');

  return [
    `📊 VPS週次レポート (${period})`,
    ``,
    `【エスタマ】${ESTAMA_HOST}`,
    block(estama),
    ``,
    `【KIRAKU】46.225.82.240`,
    block(kiraku),
    ``,
    `🟢 監視グループに自動配信 (C-060拡張)`,
  ].join('\n');
}

function sendLine(text) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      to: LINE_GROUP,
      messages: [{ type: 'text', text: text.slice(0, 4900) }],
    });
    const req = https.request(
      {
        hostname: 'api.line.me',
        path: '/v2/bot/message/push',
        method: 'POST',
        headers: {
          Authorization: 'Bearer ' + LINE_TOKEN,
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(body),
        },
      },
      (res) => {
        let data = '';
        res.on('data', (c) => (data += c));
        res.on('end', () => resolve({ status: res.statusCode, body: data }));
      }
    );
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

(async () => {
  if (!LINE_TOKEN || !LINE_GROUP) {
    console.error('FATAL: LINE_CHANNEL_ACCESS_TOKEN または LINE_MONITOR_GROUP_ID が未設定');
    process.exit(2);
  }

  const kiraku = collect('KIRAKU', sh);
  const estama = collect('エスタマ', shRemote);

  const msg = buildMessage(kiraku, estama);
  console.log(msg);

  if (process.argv.includes('--dry-run')) {
    console.log('\n[DRY-RUN] LINE送信スキップ');
    return;
  }

  const r = await sendLine(msg);
  console.log(`\nLINE送信 status=${r.status} body=${r.body}`);
  if (r.status !== 200) process.exit(1);
})().catch((e) => {
  console.error('FATAL:', e && e.stack ? e.stack : e);
  process.exit(1);
});
