/**
 * invoice_sender.js
 * KIRAKU 請求書一斉送信エンドポイント（GASからの呼び出しを受けてLINE Pushする）
 *
 * 認証: Bearer <INVOICE_SENDER_SECRET>（.envに設定）
 * 依存: @line/bot-sdk（既存）, LINE_CHANNEL_ACCESS_TOKEN（既存）
 *
 * 新規追加日: 2026-05-27（C-025）
 */

const express = require('express');
const router = express.Router();

const ENDPOINT = '/api/send-invoice';
const SHARED_SECRET = process.env.INVOICE_SENDER_SECRET || '';

function logTag(...args) {
  console.log('[invoice_sender]', ...args);
}

function authMiddleware(req, res, next) {
  if (!SHARED_SECRET) {
    return res.status(503).json({ success: false, error: 'INVOICE_SENDER_SECRET not configured' });
  }
  const auth = req.get('Authorization') || '';
  const token = auth.startsWith('Bearer ') ? auth.slice(7).trim() : null;
  if (!token || token !== SHARED_SECRET) {
    logTag('Unauthorized request from', req.ip);
    return res.status(401).json({ success: false, error: 'Unauthorized' });
  }
  next();
}

function yenFormat(n) {
  const num = Number(n);
  if (!Number.isFinite(num)) return '¥0';
  return '¥' + num.toLocaleString('ja-JP');
}

function buildMessage({ storeName, amount, dueDate, pdfUrl, payPageUrl, invoiceMonth }) {
  const lines = [
    `【KIRAKU】${invoiceMonth || ''}のご請求書をお送りします`,
    ``,
    `${storeName} 様`,
    ``,
    `ご請求金額: ${yenFormat(amount)}（税込）`,
    `お支払期限: ${dueDate || '別途ご案内'}`,
    ``,
    `■ 請求書PDF`,
    pdfUrl,
    ``,
    `■ お支払いページ（カード/銀行振込）`,
    payPageUrl,
    ``,
    `ご不明な点がございましたら、本メッセージへの返信、または以下までご連絡ください。`,
    `Mail: ec.product@telaria.tech`,
    `Tel: 082-909-2441`
  ];
  return lines.join('\n');
}

router.post(ENDPOINT, express.json({ limit: '512kb' }), authMiddleware, async (req, res) => {
  try {
    const {
      storeName,
      contractId,
      userIds,
      amount,
      dueDate,
      pdfUrl,
      payPageUrl,
      invoiceMonth,
      dryRun
    } = req.body || {};

    if (!storeName || !contractId) {
      return res.status(400).json({ success: false, error: 'storeName, contractId are required' });
    }
    if (!Array.isArray(userIds) || userIds.length === 0) {
      return res.status(400).json({ success: false, error: 'userIds must be a non-empty array' });
    }
    if (typeof amount !== 'number' || amount < 0) {
      return res.status(400).json({ success: false, error: 'amount must be a non-negative number' });
    }
    if (!pdfUrl || !payPageUrl) {
      return res.status(400).json({ success: false, error: 'pdfUrl and payPageUrl are required' });
    }

    const text = buildMessage({ storeName, amount, dueDate, pdfUrl, payPageUrl, invoiceMonth });

    if (dryRun) {
      logTag(`dryRun ${contractId} ${storeName} ${userIds.length} recipients`);
      return res.json({
        success: true,
        dryRun: true,
        contractId,
        storeName,
        recipients: userIds.length,
        previewMessage: text
      });
    }

    if (!process.env.LINE_CHANNEL_ACCESS_TOKEN) {
      return res.status(503).json({ success: false, error: 'LINE_CHANNEL_ACCESS_TOKEN not configured' });
    }

    const { messagingApi } = require('@line/bot-sdk');
    const client = new messagingApi.MessagingApiClient({
      channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN
    });

    let sent = 0;
    let failed = 0;
    const errors = [];

    for (const userId of userIds) {
      try {
        await client.pushMessage({
          to: userId,
          messages: [{ type: 'text', text }]
        });
        sent++;
        logTag(`push OK ${contractId} → ${userId.slice(0, 10)}...`);
      } catch (e) {
        failed++;
        const msg = (e && e.message) ? e.message : String(e);
        errors.push({ userId: userId.slice(0, 10) + '...', error: msg.slice(0, 200) });
        logTag(`push FAIL ${contractId} → ${userId.slice(0, 10)}...: ${msg}`);
      }
    }

    res.json({
      success: failed === 0,
      contractId,
      storeName,
      sent,
      failed,
      errors: errors.length ? errors : undefined
    });
  } catch (e) {
    logTag('Internal error:', e.message);
    res.status(500).json({
      success: false,
      error: 'Internal error',
      detail: e.message,
      stack: (e.stack || '').split('\n').slice(0, 5)
    });
  }
});

module.exports = router;
