'use strict';
/**
 * 2次代理店 入稿フォーム 受付ハンドラ
 * POST /api/agency-intake → Google Sheets台帳追記 + Gmail二重送信
 *
 * 必要な環境変数:
 *   GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET / GOOGLE_REFRESH_TOKEN  （既存流用）
 *   AGENCY_INTAKE_SHEET_ID / AGENCY_INTAKE_SHEET_TAB
 *   AGENCY_GMAIL_USER / AGENCY_GMAIL_APP_PASSWORD
 *   AGENCY_NOTIFY_SELF / AGENCY_NOTIFY_PRIMARY
 */
const { google } = require('googleapis');
const nodemailer = require('nodemailer');

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const PLANS = ['限定', '通常'];

const REQUIRED = [
  ['agency_name', '入稿代理店名'], ['agency_rep', '代理店担当者名'],
  ['agency_contact', '代理店担当者連絡先'], ['agency_email', '代理店担当者メール'],
  ['store_name', '店名'], ['pref', 'エリア(都道府県)'],
  ['store_phone', '店舗電話番号'], ['estama_id', 'エステ魂ID'],
  ['estama_pw', 'エステ魂PW'], ['plan', '契約プラン'],
  ['update_hours', '更新稼働時間'],
];

function validateIntake(body) {
  const b = body || {};
  for (const [key, label] of REQUIRED) {
    if (!b[key] || !String(b[key]).trim()) {
      return { ok: false, error: `${label}は必須です。` };
    }
  }
  if (!EMAIL_RE.test(String(b.agency_email).trim())) {
    return { ok: false, error: '代理店担当者メールアドレスの形式が正しくありません。' };
  }
  if (b.store_email && String(b.store_email).trim() && !EMAIL_RE.test(String(b.store_email).trim())) {
    return { ok: false, error: '店舗メールアドレスの形式が正しくありません。' };
  }
  if (!PLANS.includes(String(b.plan).trim())) {
    return { ok: false, error: '契約プランは「限定」または「通常」を選択してください。' };
  }
  return { ok: true };
}

function sanitizeLine(str) {
  return String(str == null ? '' : str)
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/[\x00-\x1f\x7f]/g, '')
    .trim();
}

function buildRow(d) {
  return [
    d.timestamp || '',
    sanitizeLine(d.agency_name), sanitizeLine(d.agency_rep),
    sanitizeLine(d.agency_contact), sanitizeLine(d.agency_email),
    sanitizeLine(d.store_name), sanitizeLine(d.pref), sanitizeLine(d.area),
    sanitizeLine(d.store_phone), sanitizeLine(d.store_email),
    sanitizeLine(d.estama_id), sanitizeLine(d.estama_pw),
    sanitizeLine(d.plan), sanitizeLine(d.update_hours),
    sanitizeLine(d.promo_symbol), sanitizeLine(d.recruit_pref),
    sanitizeLine(d.ip),
  ];
}

function buildEmailText(d) {
  return [
    'エステ魂 自動更新ツール 新規入稿を受け付けました。',
    '',
    `■受付日時: ${d.timestamp || ''}`,
    '',
    '【代理店情報】',
    `入稿代理店名: ${sanitizeLine(d.agency_name)}`,
    `担当者名: ${sanitizeLine(d.agency_rep)}`,
    `連絡先: ${sanitizeLine(d.agency_contact)}`,
    `メール: ${sanitizeLine(d.agency_email)}`,
    '',
    '【店舗情報】',
    `店名: ${sanitizeLine(d.store_name)}`,
    `エリア(都道府県): ${sanitizeLine(d.pref)}`,
    `エリア(小エリア): ${sanitizeLine(d.area)}`,
    `店舗電話番号: ${sanitizeLine(d.store_phone)}`,
    `店舗メール: ${sanitizeLine(d.store_email)}`,
    '',
    '【エステ魂管理情報】',
    `ID: ${sanitizeLine(d.estama_id)}`,
    `PW: ${sanitizeLine(d.estama_pw)}`,
    '',
    '【契約・更新設定】',
    `契約プラン: ${sanitizeLine(d.plan)}`,
    `更新稼働時間: ${sanitizeLine(d.update_hours)}`,
    `リアルタイム集客更新用記号: ${sanitizeLine(d.promo_symbol)}`,
    `リアルタイム求人更新用希望: ${sanitizeLine(d.recruit_pref)}`,
    '',
    `送信元IP: ${sanitizeLine(d.ip)}`,
  ].join('\n');
}

// ─── Google Sheets 台帳追記（OAuth2・既存認証を流用）───
async function saveIntakeToSheet(row) {
  const sheetId = process.env.AGENCY_INTAKE_SHEET_ID;
  const tab = process.env.AGENCY_INTAKE_SHEET_TAB || '入稿台帳';
  if (!sheetId) {
    console.warn('[agency-intake] AGENCY_INTAKE_SHEET_ID 未設定のため台帳保存をスキップ');
    return null;
  }
  const auth = new google.auth.OAuth2(process.env.GOOGLE_CLIENT_ID, process.env.GOOGLE_CLIENT_SECRET);
  auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  const sheets = google.sheets({ version: 'v4', auth });
  return await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: `${tab}!A:Q`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] },
  });
}

// ─── Gmail SMTP 送信（自社・1次代理店へ個別送信）───
function makeMailTransport() {
  return nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: 465,
    secure: true,
    auth: {
      user: process.env.AGENCY_GMAIL_USER,
      pass: process.env.AGENCY_GMAIL_APP_PASSWORD,
    },
  });
}

/**
 * 自社・1次代理店へ個別送信。
 * @returns {{ self: boolean, primary: boolean }} 各宛先の送信成否
 */
async function sendIntakeMails(d, transport = makeMailTransport()) {
  const subject = sanitizeLine(`【エステ魂 新規入稿】${d.store_name}（${d.agency_name}）`);
  const text = buildEmailText(d);
  const from = process.env.AGENCY_GMAIL_USER;
  const targets = [
    ['self', process.env.AGENCY_NOTIFY_SELF],
    ['primary', process.env.AGENCY_NOTIFY_PRIMARY],
  ];
  const result = { self: false, primary: false };
  for (const [key, to] of targets) {
    if (!to) { console.warn(`[agency-intake] ${key} 宛先未設定`); continue; }
    try {
      await transport.sendMail({ from, to, subject, text });
      result[key] = true;
    } catch (e) {
      console.error(`[agency-intake] ${key}(${to}) 送信失敗:`, e.message);
    }
  }
  return result;
}

// ─── ルート登録（副作用の入口）───
function registerAgencyIntakeRoute(app) {
  app.post('/api/agency-intake', async (req, res) => {
    try {
      const body = req.body || {};
      // ハニーポット: bot は website 項目を埋めがち。値があれば正常を装い破棄。
      if (body.website && String(body.website).trim()) {
        return res.json({ ok: true });
      }
      const v = validateIntake(body);
      if (!v.ok) return res.status(400).json({ ok: false, error: v.error });

      const ip = (req.headers['x-forwarded-for'] || req.socket?.remoteAddress || '').split(',')[0].trim();
      const timestamp = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
      const data = { ...body, ip, timestamp };

      // ① 台帳保存（失敗してもメール送信は継続）
      try {
        await saveIntakeToSheet(buildRow(data));
      } catch (e) {
        console.error('[agency-intake] 台帳保存失敗（継続）:', e.message);
      }

      // ② メール送信
      const mail = await sendIntakeMails(data);
      if (!mail.self && !mail.primary) {
        return res.status(500).json({ ok: false, error: '通知送信に失敗しました。時間をおいて再度お試しください。' });
      }
      return res.json({ ok: true });
    } catch (err) {
      console.error('[agency-intake] Error:', err);
      return res.status(500).json({ ok: false, error: 'サーバーエラーが発生しました。' });
    }
  });
}

module.exports = { registerAgencyIntakeRoute, validateIntake, sanitizeLine, buildRow, buildEmailText, saveIntakeToSheet, sendIntakeMails };
