'use strict';
// 台帳スプレッドシートを新規作成し、ヘッダ行を投入する使い捨てスクリプト。
// 実行: cd /app/chatbot && node scripts/create_agency_intake_sheet.js
require('dotenv').config();
const { google } = require('googleapis');

const HEADERS = [
  '受付日時(JST)', '入稿代理店名', '代理店担当者名', '代理店担当者連絡先',
  '代理店担当者メール', '店名', 'エリア(都道府県)', 'エリア(小エリア)',
  '店舗電話番号', '店舗メールアドレス', 'エステ魂ID', 'エステ魂PW',
  '契約プラン', '更新稼働時間', 'リアルタイム集客更新用記号',
  'リアルタイム求人更新用希望', '送信元IP',
];
const TAB = process.env.AGENCY_INTAKE_SHEET_TAB || '入稿台帳';

(async () => {
  const auth = new google.auth.OAuth2(process.env.GOOGLE_CLIENT_ID, process.env.GOOGLE_CLIENT_SECRET);
  auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  const sheets = google.sheets({ version: 'v4', auth });

  const created = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: '代理店入稿台帳' },
      sheets: [{ properties: { title: TAB } }],
    },
  });
  const id = created.data.spreadsheetId;

  await sheets.spreadsheets.values.update({
    spreadsheetId: id,
    range: `${TAB}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: [HEADERS] },
  });

  console.log('作成完了');
  console.log('AGENCY_INTAKE_SHEET_ID=' + id);
  console.log('URL: https://docs.google.com/spreadsheets/d/' + id + '/edit');
})().catch(e => { console.error('失敗:', e.message); process.exit(1); });
