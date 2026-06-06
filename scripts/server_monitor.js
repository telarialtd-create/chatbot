// C-058 サーバー監視 / GitHub Actions cron から呼ばれる
// /healthz を叩き、異常があれば監視LINEグループに通知する。
// env: SERVER_URL, HEALTHZ_SECRET, LINE_CHANNEL_ACCESS_TOKEN, LINE_MONITOR_GROUP_ID

const https = require('https');
const { URL } = require('url');

const SERVER_URL = process.env.SERVER_URL;
const HEALTHZ_SECRET = process.env.HEALTHZ_SECRET;
const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const LINE_GROUP = process.env.LINE_MONITOR_GROUP_ID;

function fetchHealthz() {
  return new Promise((resolve) => {
    const u = new URL(`${SERVER_URL}/healthz?token=${HEALTHZ_SECRET}`);
    const req = https.get(u, { timeout: 15000 }, (res) => {
      let data = '';
      res.on('data', (c) => (data += c));
      res.on('end', () => {
        try {
          resolve({ status: res.statusCode, body: JSON.parse(data) });
        } catch (_) {
          resolve({ status: res.statusCode, body: data });
        }
      });
    });
    req.on('error', (e) => resolve({ status: 0, error: e.message }));
    req.on('timeout', () => {
      req.destroy();
      resolve({ status: 0, error: 'timeout' });
    });
  });
}

// 経路の一過性揺らぎ（GitHub Actions runner⇔Hetzner欧州VPS間で発生）を
// 「VPSダウン」と誤判定しないよう、1回目失敗時は30秒待ってリトライ。
// 2回連続失敗で初めて異常とみなす。2026-06-06 凪追加（C-060改善）。
async function fetchHealthzWithRetry() {
  const first = await fetchHealthz();
  if (first.status === 200) return first;

  console.log(
    `[monitor] 1回目失敗 status=${first.status} error=${first.error || ''} → 30秒後リトライ`
  );
  await new Promise((r) => setTimeout(r, 30000));
  const second = await fetchHealthz();

  if (second.status === 200) {
    console.log('[monitor] 2回目成功 → 1回目は一過性揺らぎ判定で通知抑制');
    return second;
  }

  return {
    status: second.status,
    error: second.error,
    body: second.body,
    first_attempt: { status: first.status, error: first.error || null },
  };
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

function detectIssues(h) {
  const issues = [];
  (h.pm2 || []).forEach((p) => {
    if (p.status && p.status !== 'online') {
      issues.push(`pm2 [${p.name}] status=${p.status} restart=${p.restart_count}`);
    } else if (p.uptime_sec != null && p.uptime_sec < 120 && p.restart_count > 0) {
      issues.push(`pm2 [${p.name}] 直近再起動 uptime=${p.uptime_sec}s restart_total=${p.restart_count}`);
    }
  });
  if (h.disk_percent != null && h.disk_percent > 90) {
    issues.push(`disk使用率 ${h.disk_percent}%`);
  }
  if (h.memory_percent != null && h.memory_percent > 90) {
    issues.push(`memory使用率 ${h.memory_percent}%`);
  }
  if (h.estama && h.estama.age_sec !== null && h.estama.age_sec > 300) {
    issues.push(`エスタマheartbeat ${Math.floor(h.estama.age_sec / 60)}分途絶 (最終: ${h.estama.last_seen})`);
  }
  return issues;
}

(async () => {
  if (!SERVER_URL || !HEALTHZ_SECRET || !LINE_TOKEN || !LINE_GROUP) {
    console.error('FATAL: 必須envが未設定です');
    process.exit(2);
  }

  const r = await fetchHealthzWithRetry();
  const ts = new Date().toISOString();

  if (r.status !== 200) {
    const firstLine = r.first_attempt
      ? `1回目: status=${r.first_attempt.status} error=${r.first_attempt.error || ''}\n` +
        `2回目: status=${r.status} error=${r.error || ''}`
      : `status: ${r.status}\nerror: ${r.error || String(r.body).slice(0, 200)}`;
    const msg =
      `🚨 [KIRAKU] /healthz 到達不能（2回連続失敗）\n` +
      `時刻: ${ts}\n` +
      `${firstLine}\n` +
      `推測: KIRAKU VPS自体ダウン or chatbot停止\n` +
      `対処: Hetzner Console確認 / ssh root@46.225.82.240`;
    console.log(msg);
    await sendLine(msg);
    return;
  }

  const h = r.body;
  const issues = detectIssues(h);

  if (issues.length === 0) {
    console.log(
      JSON.stringify({
        ok: true,
        ts,
        pm2: (h.pm2 || []).map((p) => `${p.name}:${p.status}`),
        disk: h.disk_percent,
        mem: h.memory_percent,
        estama_age_sec: h.estama && h.estama.age_sec,
      })
    );
    return;
  }

  const lines = [
    `🚨 サーバー監視 異常検知`,
    `時刻: ${ts}`,
    ``,
    ...issues.map((i) => `• ${i}`),
    ``,
    `参考: disk ${h.disk_percent}% / mem ${h.memory_percent}%`,
    `pm2: ${(h.pm2 || []).map((p) => `${p.name}(${p.status})`).join(', ')}`,
    `estama age: ${h.estama && h.estama.age_sec !== null ? h.estama.age_sec + 's' : 'unknown'}`,
  ];
  const msg = lines.join('\n');
  console.log(msg);
  await sendLine(msg);
})().catch((e) => {
  console.error('FATAL:', e && e.stack ? e.stack : e);
  process.exit(1);
});
