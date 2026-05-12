// C-058 サーバー監視（healthz + heartbeat）2026-05-12
// GitHub Actions cron が /healthz を5分ごとに叩き、異常時はLINEに通知。
// エスタマVPSは毎分 /heartbeat を叩いて生存報告。

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const HEARTBEAT_DIR = '/app/chatbot/_monitor';
const HEARTBEAT_FILES = {
  estama: path.join(HEARTBEAT_DIR, 'heartbeat_estama.txt'),
};

try { fs.mkdirSync(HEARTBEAT_DIR, { recursive: true }); } catch (_) {}

function requireToken(req, secret) {
  const t = req.query?.token || req.body?.token || req.headers?.['x-monitor-token'];
  return Boolean(secret) && t === secret;
}

function getPm2List() {
  try {
    const raw = execSync('pm2 jlist 2>/dev/null', { timeout: 5000 }).toString();
    const procs = JSON.parse(raw);
    return procs.map(p => ({
      name: p.name,
      status: p.pm2_env?.status,
      restart_count: p.pm2_env?.restart_time || 0,
      cpu: p.monit?.cpu,
      memory_mb: Math.round((p.monit?.memory || 0) / 1024 / 1024),
      uptime_sec: p.pm2_env?.pm_uptime
        ? Math.floor((Date.now() - p.pm2_env.pm_uptime) / 1000)
        : 0,
    }));
  } catch (e) {
    return [{ error: e.message }];
  }
}

function getDiskPercent() {
  try {
    const out = execSync("df / | tail -1 | awk '{print $5}' | tr -d '%'", { timeout: 3000 })
      .toString().trim();
    return parseInt(out, 10);
  } catch (_) {
    return -1;
  }
}

function getMemoryPercent() {
  try {
    const out = execSync("free | awk '/^Mem:/ {printf \"%.0f\", $3/$2*100}'", { timeout: 3000 })
      .toString().trim();
    return parseInt(out, 10);
  } catch (_) {
    return -1;
  }
}

function getEstamaHeartbeat() {
  try {
    const content = fs.readFileSync(HEARTBEAT_FILES.estama, 'utf8').trim();
    const ageSec = Math.floor((Date.now() - new Date(content).getTime()) / 1000);
    return { last_seen: content, age_sec: ageSec };
  } catch (_) {
    return { last_seen: null, age_sec: null };
  }
}

function register(app) {
  // /heartbeat (GET/POST): エスタマVPSからの生存報告を受信
  const handleHeartbeat = (req, res) => {
    if (!requireToken(req, process.env.HEARTBEAT_SECRET)) {
      return res.status(401).json({ ok: false, error: 'unauthorized' });
    }
    const from = String(req.query?.from || req.body?.from || '').slice(0, 20);
    const file = HEARTBEAT_FILES[from];
    if (!file) {
      return res.status(400).json({ ok: false, error: 'unknown source' });
    }
    fs.writeFileSync(file, new Date().toISOString());
    res.json({ ok: true, from, ts: new Date().toISOString() });
  };
  app.get('/heartbeat', handleHeartbeat);
  app.post('/heartbeat', handleHeartbeat);

  // /healthz (GET): GitHub Actions から監視用に叩く
  app.get('/healthz', (req, res) => {
    if (!requireToken(req, process.env.HEALTHZ_SECRET)) {
      return res.status(401).json({ ok: false, error: 'unauthorized' });
    }
    try {
      res.json({
        ok: true,
        host: 'kiraku',
        ts: new Date().toISOString(),
        pm2: getPm2List(),
        disk_percent: getDiskPercent(),
        memory_percent: getMemoryPercent(),
        estama: getEstamaHeartbeat(),
      });
    } catch (e) {
      res.status(500).json({ ok: false, error: e.message });
    }
  });
}

module.exports = { register };
