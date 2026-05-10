const { createAuthClient } = require('/app/chatbot/lib/common');
const { google } = require('googleapis');
const fs = require('fs');
(async () => {
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });

  const TARGETS = [];
  for (let d = 3; d <= 31; d++) {
    const name = '2026年5月' + d + '日';
    const res = await drive.files.list({
      q: "name = '" + name + "' and trashed=false and mimeType='application/vnd.google-apps.spreadsheet'",
      fields: 'files(id,name,createdTime,parents,owners(emailAddress))',
      pageSize: 30,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true
    });
    res.data.files.forEach(f => TARGETS.push(f));
  }

  console.log('ゴミ箱移動対象:', TARGETS.length, '件');
  TARGETS.forEach(f => console.log('  - ' + f.name + ' | id=' + f.id + ' | created=' + f.createdTime + ' | owner=' + (f.owners||[]).map(o=>o.emailAddress).join(',')));

  // 記録ログ
  const ts = new Date().toISOString().replace(/[:.]/g, '-');
  fs.writeFileSync('/app/chatbot/_gas_backup/trashed_future_nippo_' + ts + '.json', JSON.stringify({savedAt:new Date().toISOString(), targets: TARGETS}, null, 2));

  console.log('\n=== ゴミ箱移動実行 ===');
  let okCnt = 0, ngCnt = 0;
  for (const f of TARGETS) {
    try {
      await drive.files.update({ fileId: f.id, requestBody: { trashed: true }, supportsAllDrives: true });
      console.log('  ✓ ゴミ箱: ' + f.name + ' (' + f.id + ')');
      okCnt++;
    } catch (e) {
      console.log('  ✗ 失敗: ' + f.name + ' (' + f.id + ') -> ' + e.message);
      ngCnt++;
    }
  }
  console.log('\n結果: 成功 ' + okCnt + '件 / 失敗 ' + ngCnt + '件');

  // 検証
  console.log('\n=== 検証: ゴミ箱外で5月3日〜31日の未来日報が残っていないか ===');
  let remain = 0;
  for (let d = 3; d <= 31; d++) {
    const name = '2026年5月' + d + '日';
    const res = await drive.files.list({
      q: "name = '" + name + "' and trashed=false and mimeType='application/vnd.google-apps.spreadsheet'",
      fields: 'files(id,name)',
      pageSize: 30,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true
    });
    if (res.data.files.length > 0) {
      remain += res.data.files.length;
      res.data.files.forEach(f => console.log('  残: ' + f.name + ' | ' + f.id));
    }
  }
  console.log('未来日報の残存:', remain, '件');
})().catch(e => { console.error('FATAL:', e.message); process.exit(1); });
