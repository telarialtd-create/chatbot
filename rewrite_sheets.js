const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1LQ_cFXAtaac8sZ2XsZmq3H7ZhRodCSmeneW3gzqIYBc';

let _authClient = null;
function createAuthClient() {
  if (_authClient) return _authClient;

  let client_id, client_secret, refresh_token, access_token;

  if (process.env.GOOGLE_CLIENT_ID) {
    client_id = process.env.GOOGLE_CLIENT_ID;
    client_secret = process.env.GOOGLE_CLIENT_SECRET;
    refresh_token = process.env.GOOGLE_REFRESH_TOKEN;
    access_token = process.env.GOOGLE_ACCESS_TOKEN || null;
  } else {
    const oauthKeys = JSON.parse(fs.readFileSync(path.join(process.env.HOME, '.config/gcp-oauth.keys.json')));
    const credentials = JSON.parse(fs.readFileSync(path.join(process.env.HOME, '.config/gdrive-server-credentials.json')));
    client_id = oauthKeys.installed.client_id;
    client_secret = oauthKeys.installed.client_secret;
    refresh_token = credentials.refresh_token;
    access_token = credentials.access_token;
  }

  const oauth2Client = new google.auth.OAuth2(client_id, client_secret, 'http://localhost');
  oauth2Client.setCredentials({ access_token, refresh_token });
  _authClient = oauth2Client;
  return oauth2Client;
}

async function getSheetNameByGid(sheets, gid) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheet = meta.data.sheets.find(s => s.properties.sheetId === gid);
  if (!sheet) throw new Error(`シートが見つかりません (gid=${gid})`);
  return sheet.properties.title;
}

async function clearAndWrite(sheets, sheetName, clearRange, data) {
  // Clear first
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!${clearRange}`,
  });

  // Write new data
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: data },
  });
}

// ============================================================
// SHEET 1: 現金 (gid=1180302499)
// ============================================================
async function rewriteSheet1(sheets) {
  const sheetName = await getSheetNameByGid(sheets, 1180302499);
  console.log(`Sheet 1 (現金): "${sheetName}" を書き換え中...`);

  const data = [
    ['【現金売上】月次集計'],
    ['日付', 'C 売上', 'F 売上', '合計'],
    [1, 95100, 29000, '=SUM(B3:D3)'],
    [2, 22400, '', '=SUM(B4:D4)'],
    [3, 33600, 13750, '=SUM(B5:D5)'],
    [4, 86450, 3550, '=SUM(B6:D6)'],
    [5, 62650, 14450, '=SUM(B7:D7)'],
    [6, 80000, 37100, '=SUM(B8:D8)'],
    [7, 18600, 26200, '=SUM(B9:D9)'],
    [8, 91650, 47050, '=SUM(B10:D10)'],
    [9, 51050, 44300, '=SUM(B11:D11)'],
    [10, 45500, 1450, '=SUM(B12:D12)'],
    [11, 71100, 11400, '=SUM(B13:D13)'],
    [12, 60400, 40800, '=SUM(B14:D14)'],
    [13, 66500, 72550, '=SUM(B15:D15)'],
    [14, 80400, 46500, '=SUM(B16:D16)'],
    [15, 39900, 15050, '=SUM(B17:D17)'],
    [16, 37700, 44900, '=SUM(B18:D18)'],
    [17, 13500, 26650, '=SUM(B19:D19)'],
    [18, 46950, 27800, '=SUM(B20:D20)'],
    [19, 44600, 78950, '=SUM(B21:D21)'],
    [20, 73450, 21750, '=SUM(B22:D22)'],
    [21, 33400, 34700, '=SUM(B23:D23)'],
    [22, 60450, 9650, '=SUM(B24:D24)'],
    [23, 2800, 34700, '=SUM(B25:D25)'],
    [24, 69000, 21350, '=SUM(B26:D26)'],
    [25, 25700, 31350, '=SUM(B27:D27)'],
    [26, 90900, 23350, '=SUM(B28:D28)'],
    [27, '', '', '=SUM(B29:D29)'],
    [28, 79450, 74050, '=SUM(B30:D30)'],
    [29, '', '', '=SUM(B31:D31)'],
    [30, '', '', '=SUM(B32:D32)'],
    [31, '', '', '=SUM(B33:D33)'],
    ['月合計', '=SUM(B3:B33)', '=SUM(C3:C33)', '=SUM(D3:D33)'],
  ];

  await clearAndWrite(sheets, sheetName, 'A1:J37', data);
  console.log(`  -> 完了 (${data.length}行)`);
}

// ============================================================
// SHEET 2: 支払いタスク (gid=649152626)
// ============================================================
async function rewriteSheet2(sheets) {
  const sheetName = await getSheetNameByGid(sheets, 649152626);
  console.log(`Sheet 2 (支払いタスク): "${sheetName}" を書き換え中...`);

  const data = [
    ['【定期タスク】（毎月実施）'],
    ['日付', 'タスク内容'],
    ['1日', 'テラリア領収まとめる'],
    ['1日', 'るなあゆバンス回収確認'],
    ['5日', 'エステの締め'],
    ['5日', '税理士データ送信'],
    ['10日', '給料渡し'],
    ['15日', 'まいかずテラリアからの振り込み'],
    ['15日', '税理士振り込み（封筒が届き次第）'],
    [],
    ['【月次タスク】（22日までに完了）'],
    ['No.', 'タスク内容'],
    [1, '引き落とし作成（広銀生協考慮）'],
    [2, '駐車場振り込み（広銀）'],
    [3, 'コンビニ入金（GMO）'],
    [4, 'コンビニ入金（PayPay銀行）'],
    [5, 'GMOから金を転送する'],
    [6, 'バンス回収確認打ち込み'],
    [7, 'テラリアから個人事業に振り込み'],
    [8, '社会保険 ペイジー払い'],
    [9, '市信用振り込み 131,000円'],
    [10, '事務所振り込み'],
    [],
    ['【ルーム支払い一覧】'],
    ['部屋番号', '金額（円）', '支払方法'],
    ['403', 65550, '現金振り込み'],
    ['201', 65660, '現金振り込み'],
    ['401', 64660, '現金振り込み'],
    ['601', 66550, 'クレア口座振替'],
    ['203', 63550, 'クレア口座振替'],
    ['404', 65330, '楽天銀行'],
    ['802', 67180, '楽天銀行'],
    ['1002', 69180, '楽天銀行'],
    ['801.602', 141304, 'JCBカード'],
    ['合計', '=SUM(B26:B34)', ''],
  ];

  await clearAndWrite(sheets, sheetName, 'A1:H40', data);
  console.log(`  -> 完了 (${data.length}行)`);
}

// ============================================================
// SHEET 3: G (gid=1631646919)
// ============================================================
async function rewriteSheet3(sheets) {
  const sheetName = await getSheetNameByGid(sheets, 1631646919);
  console.log(`Sheet 3 (G): "${sheetName}" を書き換え中...`);

  const data = [
    ['【G 個人収支管理】'],
    [],
    ['【収入 - SB売上】'],
    ['月', '金額（円）'],
    ['1月SB', 208440],
    ['2月SB', 95160],
    ['3月SB', 74120],
    ['4月SB', 59860],
    ['5月SB', 36040],
    ['7・8月', 100000],
    ['収入合計', '=SUM(B5:B10)'],
    [],
    ['【支出】'],
    ['項目', '金額（円）'],
    ['▲1月末', -50000],
    ['送迎車 1月', -60000],
    ['▲2月末', -50000],
    ['送迎車 2月', -60000],
    ['▲3月末', -50000],
    ['送迎車 3月', -60000],
    ['▲4月末', -50000],
    ['カード立替分M', -400719],
    ['カード立替分D', -288955],
    ['ひいろバンス', -87000],
    ['送迎車2末残り', -110997],
    ['りん', -600],
    ['飲み屋払い', -1200000],
    ['支出合計', '=SUM(B15:B27)'],
    [],
    ['差し引き合計', '=SUM(B11,B28)'],
    [],
    ['【返済計算】'],
    ['項目', '値'],
    ['元金（カード立替合計）', 2088271],
    ['毎月払い', 60000],
    ['返済期間（ヶ月）', '=B34/B35'],
    ['返済期間（年）', '=B36/12'],
  ];

  await clearAndWrite(sheets, sheetName, 'A1:L35', data);
  console.log(`  -> 完了 (${data.length}行)`);
}

// ============================================================
// SHEET 4: カード (gid=49012541)
// ============================================================
async function rewriteSheet4(sheets) {
  const sheetName = await getSheetNameByGid(sheets, 49012541);
  console.log(`Sheet 4 (カード): "${sheetName}" を書き換え中...`);

  const data = [
    ['【カード残高管理】'],
    ['時期', '区分', '内容', '取引金額', '累計', '残高', '', '次月繰越'],
    ['当月初', '-', '繰越残高', '', 0, 152347, '', '=F14'],
    ['月中', '支払い', 'タオル', -230505, '=SUM(E3,D4)', '=SUM(F3,D4)'],
    ['月中', '調整', '調整', 100000, '=SUM(E4,D5)', '=SUM(F4,D5)'],
    ['月中', '支払い', 'POVO（はじめ）', '', '=SUM(E5,D6)', '=SUM(F5,D6)'],
    ['月中', '支払い', 'POVO（はじめ）', '', '=SUM(E6,D7)', '=SUM(F6,D7)'],
    ['月中', '支払い', 'グーグルドライブ', '', '=SUM(E7,D8)', '=SUM(F7,D8)'],
    ['月中', '入金', 'カード（月中）', 94457, '=SUM(E8,D9)', '=SUM(F8,D9)'],
    ['月中', '調整', '調整', -12473, '=SUM(E9,D10)', '=SUM(F9,D10)'],
    ['20日', '支払い', '駐車場', 33000, '=SUM(E10,D11)', '=SUM(F10,D11)'],
    ['月末', '入金', 'カード（月末）', 195369, '=SUM(E11,D12)', '=SUM(F11,D12)'],
    ['月末', '支払い', 'タオル', -164092, '=SUM(E12,D13)', '=SUM(F12,D13)'],
    ['月末', 'おろし', 'おろし', -140000, '=SUM(E13,D14)', '=SUM(F13,D14)'],
    [],
    ['【カード明細確認（前払い分）】'],
    ['振込日', '対象期間', 'C金額', 'B金額', '合計'],
    ['2月17日振込', '1月1〜15日', '', '', '=SUM(C18:D18)'],
    ['2月30日振込', '1月16〜30日', '', '', '=SUM(C19:D19)'],
    [],
    ['【カード差額（PayPay計算）】'],
    ['項目', 'CREA', 'ふわもこ', '合計'],
    ['PayPay', -272000, -73000, '=SUM(B23:C23)'],
    ['カード', -299000, -139000, '=SUM(B24:C24)'],
    ['グーグル', 290, '', '=SUM(B25:C25)'],
    ['POVO', 2774, '', '=SUM(B26:C26)'],
    ['POVOチャージ', 990, '', '=SUM(B27:C27)'],
    ['合計', '=SUM(B23:B27)', '=SUM(C23:C27)', '=SUM(D23:D27)'],
  ];

  await clearAndWrite(sheets, sheetName, 'A1:Q25', data);
  console.log(`  -> 完了 (${data.length}行)`);
}

// ============================================================
// SHEET 5: 支払い (gid=0) - HEADER CLEANUP ONLY
// ============================================================
async function updateSheet5(sheets) {
  const sheetName = await getSheetNameByGid(sheets, 0);
  console.log(`Sheet 5 (支払い): "${sheetName}" のヘッダーを更新中...`);

  const updates = [
    { range: `'${sheetName}'!A1`, values: [['【月次支払い管理】']] },
    { range: `'${sheetName}'!B3`, values: [['カテゴリ']] },
    { range: `'${sheetName}'!C3`, values: [['支払先・内容']] },
    { range: `'${sheetName}'!D3`, values: [['支払方法']] },
    { range: `'${sheetName}'!E3`, values: [['予定額']] },
    { range: `'${sheetName}'!F3`, values: [['実額']] },
    { range: `'${sheetName}'!G3`, values: [['手数料']] },
    { range: `'${sheetName}'!H3`, values: [['計算用']] },
    { range: `'${sheetName}'!K3`, values: [['預かり合計']] },
    { range: `'${sheetName}'!N3`, values: [['支払チェック']] },
    { range: `'${sheetName}'!O3`, values: [['集客費(CREA)']] },
    { range: `'${sheetName}'!P3`, values: [['集客費(ふわもこ)']] },
    { range: `'${sheetName}'!R3`, values: [['口座ID']] },
    { range: `'${sheetName}'!S3`, values: [['口座残高']] },
    { range: `'${sheetName}'!X3`, values: [['口座計算用']] },
    { range: `'${sheetName}'!Z3`, values: [['振込み']] },
    // Clean up garbage in rows 36-37
    { range: `'${sheetName}'!E36`, values: [['']] },
    { range: `'${sheetName}'!E37`, values: [['']] },
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: updates,
    },
  });

  console.log(`  -> 完了 (${updates.length}セル更新)`);
}

// ============================================================
// SHEET 6: 計算用 (gid=558747688) - SECTION HEADER CLEANUP ONLY
// ============================================================
async function updateSheet6(sheets) {
  const sheetName = await getSheetNameByGid(sheets, 558747688);
  console.log(`Sheet 6 (計算用): "${sheetName}" のヘッダーを更新中...`);

  const updates = [
    { range: `'${sheetName}'!C5`, values: [['【CREA 経費管理】']] },
    { range: `'${sheetName}'!M5`, values: [['【ふわもこSPA 経費管理】']] },
    { range: `'${sheetName}'!R5`, values: [['【ダリア 人件費】']] },
    { range: `'${sheetName}'!C17`, values: [['支払先']] },
    { range: `'${sheetName}'!D17`, values: [['支払内容']] },
    { range: `'${sheetName}'!E17`, values: [['金額']] },
    { range: `'${sheetName}'!F17`, values: [['備考']] },
    { range: `'${sheetName}'!H17`, values: [['支払先']] },
    { range: `'${sheetName}'!I17`, values: [['支払内容']] },
    { range: `'${sheetName}'!J17`, values: [['金額']] },
    { range: `'${sheetName}'!K17`, values: [['備考']] },
    { range: `'${sheetName}'!N17`, values: [['支払先']] },
    { range: `'${sheetName}'!O17`, values: [['支払内容']] },
    { range: `'${sheetName}'!P17`, values: [['金額']] },
    { range: `'${sheetName}'!Q17`, values: [['備考']] },
    { range: `'${sheetName}'!C30`, values: [['【固定費】']] },
    { range: `'${sheetName}'!H30`, values: [['【固定費】']] },
    { range: `'${sheetName}'!N30`, values: [['【固定費】']] },
    { range: `'${sheetName}'!C39`, values: [['【変動費】']] },
    { range: `'${sheetName}'!H39`, values: [['【変動費】']] },
    { range: `'${sheetName}'!N39`, values: [['【変動費】']] },
    { range: `'${sheetName}'!C49`, values: [['【特殊】']] },
    { range: `'${sheetName}'!H49`, values: [['【特殊】']] },
    { range: `'${sheetName}'!N49`, values: [['【特殊】']] },
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: updates,
    },
  });

  console.log(`  -> 完了 (${updates.length}セル更新)`);
}

// ============================================================
// MAIN
// ============================================================
async function main() {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const results = [];

  // Sheet 1: 現金
  try {
    await rewriteSheet1(sheets);
    results.push({ sheet: '現金 (gid=1180302499)', status: '成功' });
  } catch (e) {
    console.error(`Sheet 1 エラー:`, e.message);
    results.push({ sheet: '現金 (gid=1180302499)', status: '失敗', error: e.message });
  }

  // Sheet 2: 支払いタスク
  try {
    await rewriteSheet2(sheets);
    results.push({ sheet: '支払いタスク (gid=649152626)', status: '成功' });
  } catch (e) {
    console.error(`Sheet 2 エラー:`, e.message);
    results.push({ sheet: '支払いタスク (gid=649152626)', status: '失敗', error: e.message });
  }

  // Sheet 3: G
  try {
    await rewriteSheet3(sheets);
    results.push({ sheet: 'G (gid=1631646919)', status: '成功' });
  } catch (e) {
    console.error(`Sheet 3 エラー:`, e.message);
    results.push({ sheet: 'G (gid=1631646919)', status: '失敗', error: e.message });
  }

  // Sheet 4: カード
  try {
    await rewriteSheet4(sheets);
    results.push({ sheet: 'カード (gid=49012541)', status: '成功' });
  } catch (e) {
    console.error(`Sheet 4 エラー:`, e.message);
    results.push({ sheet: 'カード (gid=49012541)', status: '失敗', error: e.message });
  }

  // Sheet 5: 支払い (partial update)
  try {
    await updateSheet5(sheets);
    results.push({ sheet: '支払い (gid=0)', status: '成功' });
  } catch (e) {
    console.error(`Sheet 5 エラー:`, e.message);
    results.push({ sheet: '支払い (gid=0)', status: '失敗', error: e.message });
  }

  // Sheet 6: 計算用 (partial update)
  try {
    await updateSheet6(sheets);
    results.push({ sheet: '計算用 (gid=558747688)', status: '成功' });
  } catch (e) {
    console.error(`Sheet 6 エラー:`, e.message);
    results.push({ sheet: '計算用 (gid=558747688)', status: '失敗', error: e.message });
  }

  console.log('\n============================================================');
  console.log('【実行結果サマリー】');
  console.log('============================================================');
  for (const r of results) {
    const mark = r.status === '成功' ? '✓' : '✗';
    console.log(`${mark} ${r.sheet}: ${r.status}${r.error ? ' - ' + r.error : ''}`);
  }
  const success = results.filter(r => r.status === '成功').length;
  console.log(`\n合計: ${success}/${results.length} シート処理完了`);
}

main().catch(e => {
  console.error('予期せぬエラー:', e);
  process.exit(1);
});
