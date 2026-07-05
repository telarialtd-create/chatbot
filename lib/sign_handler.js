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

      const appendResp = await saveToSheet({
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

      res.json({ ok: true });

      // ─── 同意書PDFを生成しDriveへ保存（ベストエフォート・応答後に非同期実行）───
      makeAgreementPdfAndStore({
        company:      body.company.trim(),
        rep:          body.signature.trim(),
        email:        body.email.trim(),
        phone:        body.phone.trim(),
        timestamp:    agreedAt,
        ip,
        termsVersion: body.terms_version || 'ツール利用規約（版数不明）',
        stamp:        new Date().toISOString().replace(/[:T]/g, '-').slice(0, 16),
      }, appendResp)
        .then(link => console.log('[agree-tool] 同意書PDF保存:', link))
        .catch(e => console.warn('[agree-tool] PDF生成スキップ:', e.message));
      return;

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

  return await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: `${tabName}!A:O`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] },
  });
}

// ─── OAuth2 クライアント（Sheets/Drive 共通）──────────────
function makeOAuth() {
  const auth = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET
  );
  auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  return auth;
}

// ─── 同意書PDFを生成し Drive に保存、台帳P列へリンクを書き戻す ───
// ※ベストエフォート。失敗しても署名記録（台帳の行）は既に保存済みのため影響しない。
async function makeAgreementPdfAndStore(data, appendResp) {
  const html = buildAgreementHtml(data);

  // 1) puppeteer で HTML → PDF
  const puppeteer = require('puppeteer');
  const browser = await puppeteer.launch({
    headless: true,
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });
  let pdfBuffer;
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '16mm', bottom: '16mm', left: '14mm', right: '14mm' },
    });
  } finally {
    await browser.close();
  }

  // 2) Drive アップロード（専用フォルダ「電子サイン同意書」＝03_営業・販売資料 配下）
  const auth  = makeOAuth();
  const drive = google.drive({ version: 'v3', auth });
  const parent = process.env.AGREE_PDF_FOLDER_ID || '1N_MQm5s0STVUTfGt4x-AzA2H_UmLve1-';
  const safe = s => String(s || '').replace(/[\\/:*?"<>|\n\r]/g, '_').slice(0, 40);
  const fname = `同意書_${safe(data.company)}_${safe(data.rep)}_${data.stamp}.pdf`;

  const { Readable } = require('stream');
  const pdfStream = new Readable();
  pdfStream.push(Buffer.from(pdfBuffer)); // Uint8Array→Buffer（1チャンクとして流す）
  pdfStream.push(null);
  const created = await drive.files.create({
    requestBody: { name: fname, mimeType: 'application/pdf', parents: parent ? [parent] : undefined },
    media: { mimeType: 'application/pdf', body: pdfStream },
    fields: 'id, webViewLink',
    supportsAllDrives: true,
  });
  const link = created.data.webViewLink || `https://drive.google.com/file/d/${created.data.id}/view`;

  // 3) 台帳の該当行 P列へ PDF リンクを書き戻す
  try {
    const ur = appendResp?.data?.updates?.updatedRange || '';
    const m = ur.match(/![A-Z]+(\d+)/);
    if (m) {
      const tabName = process.env.SIGN_SHEET_TAB || '電子サイン台帳';
      const sheets = google.sheets({ version: 'v4', auth });
      await sheets.spreadsheets.values.update({
        spreadsheetId: process.env.SIGN_SHEET_ID,
        range: `${tabName}!P${m[1]}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[link]] },
      });
    }
  } catch (e) {
    console.warn('[agree-tool] PDFリンクの台帳書き戻しスキップ:', e.message);
  }

  return link;
}

// ─── 同意書PDFのHTML（A4・日本語Notoフォント）────────────
function buildAgreementHtml(d) {
  const esc = s => String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  return `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Noto Sans CJK JP','Noto Serif CJK JP',sans-serif; color: #111; font-size: 12px; line-height: 1.8; padding: 4px 6px; }
    .head { text-align: center; border-bottom: 2px solid #111; padding-bottom: 14px; margin-bottom: 22px; }
    .head h1 { font-size: 20px; letter-spacing: 0.08em; }
    .head .sub { font-size: 11px; color: #666; margin-top: 6px; }
    .lead { font-size: 12px; margin-bottom: 18px; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 22px; }
    th, td { border: 1px solid #bbb; padding: 8px 12px; text-align: left; vertical-align: top; font-size: 12px; }
    th { background: #f2f2f2; width: 34%; font-weight: 700; }
    .sig { font-size: 16px; font-weight: 700; letter-spacing: 0.05em; }
    .terms { border: 1px solid #ccc; background: #fafafa; border-radius: 4px; padding: 14px 16px; font-size: 10.5px; line-height: 1.85; margin-bottom: 20px; }
    .terms h2 { font-size: 12px; margin-bottom: 8px; }
    .terms p { margin-bottom: 8px; }
    .foot { margin-top: 26px; font-size: 10.5px; color: #444; border-top: 1px solid #ddd; padding-top: 12px; }
  </style></head><body>
    <div class="head">
      <h1>電子署名 同意書</h1>
      <div class="sub">ツール利用規約に対する同意の記録</div>
    </div>
    <p class="lead">下記の者は、本書に記載の日時において、株式会社TeLARIA が提供するツール利用規約（下記要旨を含む全条項）の内容を確認・理解のうえ、これに同意し、電子署名を行いました。</p>
    <table>
      <tr><th>会社名・屋号</th><td>${esc(d.company)}</td></tr>
      <tr><th>担当者氏名（電子署名）</th><td class="sig">${esc(d.rep)}</td></tr>
      <tr><th>メールアドレス</th><td>${esc(d.email)}</td></tr>
      <tr><th>電話番号</th><td>${esc(d.phone)}</td></tr>
      <tr><th>同意日時</th><td>${esc(d.timestamp)}（日本時間）</td></tr>
      <tr><th>送信元IPアドレス</th><td>${esc(d.ip)}</td></tr>
      <tr><th>同意した規約</th><td>${esc(d.termsVersion)}</td></tr>
    </table>
    <div class="terms">
      <h2>同意条項の要旨（ツール利用規約より）</h2>
      <p><strong>第5条（外部サービス利用リスク・免責）</strong>：本ツール（エスタマ自動更新等）の利用に起因して、外部サイト運営者により利用者のアカウントが停止・凍結・削除等される可能性があること、およびこれにより生じた一切の損害について当社が一切の責任を負わないことを、利用者は承諾する。</p>
      <p><strong>第6条（本ツールの不具合・停止・利用不能の免責）</strong>：当社は本ツールの継続的な正常稼働を保証せず、不具合・停止・利用不能により生じた一切の損害について責任を負わない。</p>
      <p><strong>第7条（包括免責）</strong>：本ツールの利用または利用不能に起因する一切の損害について、当社の故意・重過失がある場合を除き責任を負わない。第三者との紛争は利用者が自己の責任で解決する。</p>
      <p style="color:#666;">※本書は要旨です。規約全文は当社ウェブサイトに掲載の「ツール利用規約」によります。</p>
    </div>
    <div class="foot">
      本同意書は、利用者による電子署名（氏名の入力）および同意チェックの操作をもって成立し、同意日時・送信元IPアドレスとともに当社の電子サイン台帳に記録されています。<br>
      発行者：株式会社TeLARIA ／ ツール利用規約 電子同意システム
    </div>
  </body></html>`;
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
