'use strict';

const crypto = require('crypto');
const path   = require('path');
const { google } = require('googleapis');

// セッショントークン管理（メモリ）
const sessions = new Map(); // token → expiresAt
const SESSION_TTL = 8 * 60 * 60 * 1000; // 8時間

const VALID_STATUSES = ['未対応', '対応中', '契約済み'];

function getAuth() {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET
  );
  oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  return oauth2Client;
}

function parseCookies(header = '') {
  const result = {};
  for (const part of header.split(';')) {
    const [k, ...v] = part.trim().split('=');
    if (k) result[k.trim()] = decodeURIComponent(v.join('='));
  }
  return result;
}

function isAuthenticated(req) {
  const cookies = parseCookies(req.headers.cookie || '');
  const token = cookies['admin_token'];
  if (!token) return false;
  const exp = sessions.get(token);
  if (!exp || Date.now() > exp) {
    sessions.delete(token);
    return false;
  }
  return true;
}

function registerAdminRoutes(app) {

  // ── 管理画面HTML ──────────────────────────────────────────
  app.get('/admin', (_req, res) => {
    res.sendFile(path.join(__dirname, '../public/admin.html'));
  });

  // ── ログイン ─────────────────────────────────────────────
  app.post('/admin/login', (req, res) => {
    const { password } = req.body || {};
    if (!process.env.ADMIN_PASSWORD || password !== process.env.ADMIN_PASSWORD) {
      return res.status(401).json({ ok: false, error: 'パスワードが違います。' });
    }
    const token = crypto.randomBytes(32).toString('hex');
    sessions.set(token, Date.now() + SESSION_TTL);
    res.setHeader('Set-Cookie',
      `admin_token=${token}; HttpOnly; Secure; SameSite=Strict; Path=/admin; Max-Age=28800`
    );
    return res.json({ ok: true });
  });

  // ── ログアウト ────────────────────────────────────────────
  app.post('/admin/logout', (req, res) => {
    const cookies = parseCookies(req.headers.cookie || '');
    sessions.delete(cookies['admin_token']);
    res.setHeader('Set-Cookie', 'admin_token=; HttpOnly; Path=/admin; Max-Age=0');
    return res.json({ ok: true });
  });

  // ── 申し込み一覧取得 ──────────────────────────────────────
  app.get('/admin/api/applications', async (req, res) => {
    if (!isAuthenticated(req)) return res.status(401).json({ ok: false, error: '認証が必要です。' });

    const sheetId = process.env.SIGN_SHEET_ID;
    const tabName = process.env.SIGN_SHEET_TAB || '電子サイン台帳';
    if (!sheetId) return res.json({ ok: true, data: [] });

    try {
      const auth   = getAuth();
      const sheets = google.sheets({ version: 'v4', auth });
      const resp   = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: `${tabName}!A2:M`,
      });

      const rows = resp.data.values || [];
      const data = rows.map((row, i) => ({
        rowIndex: i + 2,
        timestamp: row[0]  || '',
        company:   row[1]  || '',
        store:     row[2]  || '',
        rep:       row[3]  || '',
        position:  row[4]  || '',
        email:     row[5]  || '',
        phone:     row[6]  || '',
        address:   row[7]  || '',
        plan:      row[8]  || '',
        signature: row[9]  || '',
        message:   row[10] || '',
        ip:        row[11] || '',
        status:    row[12] || '未対応',
      }));

      return res.json({ ok: true, data: data.reverse() }); // 新着順
    } catch (err) {
      console.error('[admin_handler] list error:', err);
      return res.status(500).json({ ok: false, error: 'データ取得に失敗しました。' });
    }
  });

  // ── ステータス更新 ────────────────────────────────────────
  app.patch('/admin/api/applications/:row/status', async (req, res) => {
    if (!isAuthenticated(req)) return res.status(401).json({ ok: false, error: '認証が必要です。' });

    const rowNum = parseInt(req.params.row, 10);
    const { status } = req.body || {};

    if (!VALID_STATUSES.includes(status)) {
      return res.status(400).json({ ok: false, error: '無効なステータスです。' });
    }
    if (isNaN(rowNum) || rowNum < 2) {
      return res.status(400).json({ ok: false, error: '行番号が不正です。' });
    }

    const sheetId = process.env.SIGN_SHEET_ID;
    const tabName = process.env.SIGN_SHEET_TAB || '電子サイン台帳';

    try {
      const auth   = getAuth();
      const sheets = google.sheets({ version: 'v4', auth });
      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: `${tabName}!M${rowNum}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[status]] },
      });
      return res.json({ ok: true });
    } catch (err) {
      console.error('[admin_handler] status update error:', err);
      return res.status(500).json({ ok: false, error: 'ステータス更新に失敗しました。' });
    }
  });
}

module.exports = { registerAdminRoutes };
