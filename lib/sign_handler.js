'use strict';

/**
 * KIRAKU 電子サイン処理モジュール
 * POST /api/sign → Google Sheets保存 + LINE通知
 *
 * 必要な環境変数:
 *   SIGN_SHEET_ID         : 同意書台帳のスプレッドシートID
 *   SIGN_SHEET_TAB        : タブ名（例: "電子サイン台帳"）
 *   GOOGLE_SERVICE_ACCOUNT_EMAIL
 *   GOOGLE_PRIVATE_KEY
 *   LINE_CHANNEL_ACCESS_TOKEN
 *   SIGN_NOTIFY_LINE_ID   : 通知先のLINE UserID or GroupID
 */

const { google } = require('googleapis');

const PLAN_LABELS = {
  starter:  'STARTER（¥9,800/月）',
  standard: 'STANDARD（¥19,800/月）',
  premium:  'PREMIUM（¥45,000/月）',
  custom:   'CUSTOM（個別見積もり）',
};

/**
 * Express app に /api/sign ルートを登録する
 * @param {import('express').Express} app
 */
function registerSignRoute(app) {
  app.post('/api/sign', async (req, res) => {
    try {
      const body = req.body || {};
      const ip = (req.headers['x-forwarded-for'] || req.socket?.remoteAddress || '').split(',')[0].trim();

      // ─── バリデーション ───
      const required = ['plan', 'company_name', 'rep_name', 'email', 'signature'];
      for (const key of required) {
        if (!body[key] || !String(body[key]).trim()) {
          return res.status(400).json({ ok: false, error: `${key} は必須です。` });
        }
      }
      if (!PLAN_LABELS[body.plan]) {
        return res.status(400).json({ ok: false, error: '無効なプランです。' });
      }
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(body.email)) {
        return res.status(400).json({ ok: false, error: 'メールアドレスの形式が正しくありません。' });
      }
      // 署名と担当者名が一致するかを確認（軽微な差異は許容）
      const sigNorm = body.signature.trim().replace(/\s+/g, '');
      const repNorm = body.rep_name.trim().replace(/\s+/g, '');
      if (sigNorm !== repNorm) {
        return res.status(400).json({ ok: false, error: '電子署名と担当者氏名が一致しません。' });
      }
      // 外部サービス利用リスクの承諾を必須とする（利用規約 第9条の2）
      if (body.risk_agreed !== true) {
        return res.status(400).json({ ok: false, error: '外部サービス利用リスク（第9条の2）への同意が必要です。' });
      }

      const agreedAt = body.agreed_at ? new Date(body.agreed_at).toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }) : new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });

      // ─── Google Sheets 保存 ───
      await saveToSheet({
        timestamp: agreedAt,
        company:   body.company_name,
        store:     body.store_name || '',
        rep:       body.rep_name,
        position:  body.position || '',
        email:     body.email,
        phone:     body.phone || '',
        address:   body.address || '',
        plan:      PLAN_LABELS[body.plan],
        signature: body.signature,
        message:   body.message || '',
        ip,
        riskAck:   body.risk_ack_text || '外部サービス利用リスク（第9条の2）に同意',
      });

      // ─── LINE 通知 ───
      await notifyLine({
        company: body.company_name,
        store:   body.store_name || '',
        rep:     body.rep_name,
        email:   body.email,
        plan:    PLAN_LABELS[body.plan],
        agreedAt,
      });

      return res.json({ ok: true });

    } catch (err) {
      console.error('[sign_handler] Error:', err);
      return res.status(500).json({ ok: false, error: 'サーバーエラーが発生しました。' });
    }
  });

  // ─── ツール利用規約 同意（料金なし・免責特化版 riyoukiyaku-tool.html 用）───
  app.post('/api/agree-tool', async (req, res) => {
    try {
      const body = req.body || {};
      const ip = (req.headers['x-forwarded-for'] || req.socket?.remoteAddress || '').split(',')[0].trim();

      if (!body.company || !String(body.company).trim()) {
        return res.status(400).json({ ok: false, error: '会社名・屋号は必須です。' });
      }
      if (!body.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(body.email).trim())) {
        return res.status(400).json({ ok: false, error: 'メールアドレスの形式が正しくありません。' });
      }
      if (!body.phone || !String(body.phone).trim()) {
        return res.status(400).json({ ok: false, error: '電話番号は必須です。' });
      }
      if (!body.signature || !String(body.signature).trim()) {
        return res.status(400).json({ ok: false, error: '担当者氏名（電子署名）は必須です。' });
      }
      if (body.agreed !== true) {
        return res.status(400).json({ ok: false, error: 'ツール利用規約への同意が必要です。' });
      }

      const agreedAt = body.agreed_at
        ? new Date(body.agreed_at).toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })
        : new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });

      await saveToSheet({
        timestamp: agreedAt,
        company:   body.company.trim(),
        store:     '',
        rep:       body.signature.trim(),
        position:  '',
        email:     body.email.trim(),
        phone:     body.phone.trim(),
        address:   '',
        plan:      'ツール利用規約 同意（料金なし）',
        signature: body.signature.trim(),
        message:   '',
        ip,
        riskAck:   'ツール利用規約（第5条=外部サービス利用リスク/第6条=ツール不保証/第7条=包括免責の全条項）に同意',
        termsVersion: body.terms_version || 'ツール利用規約（版数不明）',
      });

      try {
        await notifyLine({
          company: body.company || '(未記入)',
          store:   '',
          rep:     body.signature.trim(),
          email:   '(ツール利用規約 同意)',
          plan:    'ツール利用規約 同意',
          agreedAt,
        });
      } catch (e) {
        console.warn('[agree-tool] LINE通知スキップ:', e.message);
      }

      return res.json({ ok: true });

    } catch (err) {
      console.error('[agree-tool] Error:', err);
      return res.status(500).json({ ok: false, error: 'サーバーエラーが発生しました。' });
    }
  });
}

// ─── Google Sheets 書き込み ───────────────────────────────
async function saveToSheet(data) {
  const sheetId  = process.env.SIGN_SHEET_ID;
  const tabName  = process.env.SIGN_SHEET_TAB || '電子サイン台帳';

  if (!sheetId) {
    console.warn('[sign_handler] SIGN_SHEET_ID が未設定のため Sheets 保存をスキップ');
    return;
  }

  const auth = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET
  );
  auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });

  const sheets = google.sheets({ version: 'v4', auth });

  const row = [
    data.timestamp,
    data.company,
    data.store,
    data.rep,
    data.position,
    data.email,
    data.phone,
    data.address,
    data.plan,
    data.signature,
    data.message,
    data.ip,
    '未対応',
    data.riskAck || '',
    data.termsVersion || '',
  ];

  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: `${tabName}!A:O`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] },
  });
}

// ─── LINE 通知 ────────────────────────────────────────────
async function notifyLine(data) {
  const token  = process.env.LINE_CHANNEL_ACCESS_TOKEN;
  const toId   = process.env.SIGN_NOTIFY_LINE_ID;

  if (!token || !toId) {
    console.warn('[sign_handler] LINE通知の設定が未完了のためスキップ');
    return;
  }

  const text = [
    '📝 新しい申込みがありました',
    '',
    `プラン　: ${data.plan}`,
    `会社名　: ${data.company}${data.store ? ' / ' + data.store : ''}`,
    `担当者　: ${data.rep}`,
    `メール　: ${data.email}`,
    `同意日時: ${data.agreedAt}`,
    '',
    '→ 電子サイン台帳で詳細を確認してください。',
  ].join('\n');

  const body = JSON.stringify({
    to: toId,
    messages: [{ type: 'text', text }],
  });

  const https = require('https');
  const options = {
    hostname: 'api.line.me',
    path: '/v2/bot/message/push',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`,
      'Content-Length': Buffer.byteLength(body),
    },
  };

  await new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let raw = '';
      res.on('data', d => raw += d);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) resolve();
        else reject(new Error(`LINE API: ${res.statusCode} ${raw}`));
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

module.exports = { registerSignRoute };
