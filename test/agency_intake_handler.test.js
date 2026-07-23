'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const h = require('../lib/agency_intake_handler');

const validBody = {
  agency_name: 'ハーツパートナーズ', agency_rep: '山田太郎',
  agency_contact: '09011112222', agency_email: 'agent@example.com',
  store_name: 'テスト店', pref: '静岡県', area: '静岡市',
  store_phone: '0541234567', store_email: '',
  estama_id: 'esthe.test@icloud.com', estama_pw: 'Pw12345',
  plan: '限定', update_hours: '9:30〜翌3:00',
  promo_symbol: '★', recruit_pref: '応募資格', website: '',
};

test('validateIntake: 正常系はok=true', () => {
  assert.deepStrictEqual(h.validateIntake(validBody), { ok: true });
});

test('validateIntake: 必須欠落はok=false', () => {
  const b = { ...validBody, store_name: '' };
  const r = h.validateIntake(b);
  assert.strictEqual(r.ok, false);
  assert.match(r.error, /店名/);
});

test('validateIntake: 代理店メール形式不正はok=false', () => {
  const r = h.validateIntake({ ...validBody, agency_email: 'bad' });
  assert.strictEqual(r.ok, false);
  assert.match(r.error, /メール/);
});

test('validateIntake: 契約プランが限定/通常以外はok=false', () => {
  const r = h.validateIntake({ ...validBody, plan: 'VIP' });
  assert.strictEqual(r.ok, false);
  assert.match(r.error, /プラン/);
});

test('validateIntake: 店舗メールは任意だが入力時は形式チェック', () => {
  assert.strictEqual(h.validateIntake({ ...validBody, store_email: 'x@y.jp' }).ok, true);
  assert.strictEqual(h.validateIntake({ ...validBody, store_email: 'bad' }).ok, false);
});

test('sanitizeLine: 改行と制御文字を空白化', () => {
  // \r\n が1つの空白に、\t が1つの空白になる（連続する改行類はまとめて1空白）
  assert.strictEqual(h.sanitizeLine('あ\r\nい\tう'), 'あ い う');
});

test('buildRow: 17列・RAW書込のため電話に先頭アポストロフィは付与しない', () => {
  const row = h.buildRow({ ...validBody, timestamp: '2026-07-24 10:00:00', ip: '1.2.3.4' });
  assert.strictEqual(row.length, 17);
  assert.strictEqual(row[0], '2026-07-24 10:00:00');
  assert.strictEqual(row[8], '0541234567'); // 店舗電話
  assert.strictEqual(row[3], '09011112222'); // 代理店連絡先
  assert.strictEqual(row[16], '1.2.3.4');
});

test('buildEmailText: 全項目とID/PWを含む', () => {
  const txt = h.buildEmailText({ ...validBody, timestamp: '2026-07-24 10:00:00', ip: '1.2.3.4' });
  assert.match(txt, /テスト店/);
  assert.match(txt, /esthe\.test@icloud\.com/);
  assert.match(txt, /Pw12345/);
  assert.match(txt, /限定/);
});

test('sendIntakeMails: self成功・primary失敗を個別に返す', async () => {
  process.env.AGENCY_NOTIFY_SELF = 'self@example.com';
  process.env.AGENCY_NOTIFY_PRIMARY = 'primary@example.com';
  const fakeTransport = {
    sendMail: async ({ to }) => {
      if (to === 'primary@example.com') throw new Error('primary送信失敗（テスト用）');
      return { ok: true };
    },
  };
  const result = await h.sendIntakeMails(
    { ...validBody, timestamp: '2026-07-24 10:00:00', ip: '1.2.3.4' },
    fakeTransport
  );
  assert.deepStrictEqual(result, { self: true, primary: false });
});

test('sendIntakeMails: 両方成功', async () => {
  process.env.AGENCY_NOTIFY_SELF = 'self@example.com';
  process.env.AGENCY_NOTIFY_PRIMARY = 'primary@example.com';
  const fakeTransport = { sendMail: async () => ({ ok: true }) };
  const result = await h.sendIntakeMails(
    { ...validBody, timestamp: '2026-07-24 10:00:00', ip: '1.2.3.4' },
    fakeTransport
  );
  assert.deepStrictEqual(result, { self: true, primary: true });
});

test('sendIntakeMails: 両方失敗', async () => {
  process.env.AGENCY_NOTIFY_SELF = 'self@example.com';
  process.env.AGENCY_NOTIFY_PRIMARY = 'primary@example.com';
  const fakeTransport = { sendMail: async () => { throw new Error('送信失敗（テスト用）'); } };
  const result = await h.sendIntakeMails(
    { ...validBody, timestamp: '2026-07-24 10:00:00', ip: '1.2.3.4' },
    fakeTransport
  );
  assert.deepStrictEqual(result, { self: false, primary: false });
});

test('registerAgencyIntakeRoute: ハニーポット項目に値があれば検証・保存・送信をスキップしok=trueを返す', async () => {
  let captured;
  const fakeApp = { post: (path, handler) => { captured = { path, handler }; } };
  h.registerAgencyIntakeRoute(fakeApp);
  assert.strictEqual(captured.path, '/api/agency-intake');

  const req = {
    body: { ...validBody, website: 'http://spam.example.com' },
    headers: {},
    socket: {},
  };
  let statusCode = null;
  let jsonBody = null;
  const res = {
    status(code) { statusCode = code; return this; },
    json(body) { jsonBody = body; return this; },
  };
  await captured.handler(req, res);
  assert.strictEqual(statusCode, null); // status()未呼び出し＝デフォルト200相当
  assert.deepStrictEqual(jsonBody, { ok: true });
});
