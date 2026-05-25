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

// 1ファイルからCPU使用率（user+system+nice+iowait）の Average を取得
function cpuLineFromSar(runner, file) {
  try {
    const out = runner(`test -f ${file} && sar -u -f ${file} 2>/dev/null | awk '/^Average:/ {print 100-$NF}' || echo ''`);
    const v = parseFloat(out);
    return isFinite(v) ? v : null;
  } catch (_) {
    return null;
  }
}

function loadAvgFromSar(runner, file) {
  try {
    const out = runner(`test -f ${file} && sar -q -f ${file} 2>/dev/null | awk '/^Average:/ {print $4}' || echo ''`);
    const v = parseFloat(out);
    return isFinite(v) ? v : null;
  } catch (_) {
    return null;
  }
}

function loadPeakFromSar(runner, file) {
  try {
    const out = runner(`test -f ${file} && sar -q -f ${file} 2>/dev/null | awk '/^[0-9]/ {print $4}' | sort -n | tail -1 || echo ''`);
    const v = parseFloat(out);
    return isFinite(v) ? v : null;
  } catch (_) {
    return null;
  }
}

function memPercentFromSar(runner, file) {
  try {
    const out = runner(`test -f ${file} && sar -r -f ${file} 2>/dev/null | awk '/^Average:/ {print $5}' || echo ''`);
    const v = parseFloat(out);
    return isFinite(v) ? v : null;
  } catch (_) {
    return null;
  }
}

function cpuPeakFromSar(runner, file) {
  try {
    const out = runner(`test -f ${file} && sar -u -f ${file} 2>/dev/null | awk '/^[0-9]/ {print 100-$NF}' | sort -n | tail -1 || echo ''`);
    const v = parseFloat(out);
    return isFinite(v) ? v : null;
  } catch (_) {
    return null;
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

function collect(label, runner) {
  const files = last7SarFiles();
  const cpuAvgs = files.map((f) => cpuLineFromSar(runner, f)).filter((x) => x != null);
  const cpuPeaks = files.map((f) => cpuPeakFromSar(runner, f)).filter((x) => x != null);
  const loadAvgs = files.map((f) => loadAvgFromSar(runner, f)).filter((x) => x != null);
  const loadPeaks = files.map((f) => loadPeakFromSar(runner, f)).filter((x) => x != null);
  const memAvgs = files.map((f) => memPercentFromSar(runner, f)).filter((x) => x != null);

  const avg = (a) => (a.length ? a.reduce((s, v) => s + v, 0) / a.length : null);
  const max = (a) => (a.length ? Math.max(...a) : null);

  return {
    label,
    days: cpuAvgs.length,
    cpu_avg: avg(cpuAvgs),
    cpu_peak: max(cpuPeaks),
    load_avg: avg(loadAvgs),
    load_peak: max(loadPeaks),
    mem_avg: avg(memAvgs),
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
