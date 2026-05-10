const { createAuthClient } = require('/app/chatbot/lib/common');
const { google } = require('googleapis');
const fs = require('fs');
(async () => {
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });

  // ログから対象を再取得
  const logs = fs.readdirSync('/app/chatbot/_gas_backup/').filter(f => f.startsWith('trashed_future_nippo_'));
  if (logs.length === 0) {
    console.log('ログが見つかりません');
    return;
  }
  const latest = logs.sort().pop();
  const data = JSON.parse(fs.readFileSync('/app/chatbot/_gas_backup/' + latest, 'utf8'));
  console.log('ログ:', latest, '対象:', data.targets.length, '件');
  data.targets.forEach(f => console.log('  - ' + f.name + ' | id=' + f.id));

  // ゴミ箱移動
  console.log('\n=== ゴミ箱移動実行 ===');
  let okCnt = 0, ngCnt = 0;
  for (const f of data.targets) {
    try {
      await drive.files.update({ fileId: f.id, requestBody: { trashed: true }, supportsAllDrives: true });
      console.log('  ✓ ' + f.name + ' (' + f.id.slice(0, 12) + '...)');
      okCnt++;
    } catch (e) {
      console.log('  ✗ ' + f.name + ' (' + f.id.slice(0, 12) + '...) -> ' + e.message);
      ngCnt++;
    }
  }
  console.log('\n結果: 成功 ' + okCnt + '件 / 失敗 ' + ngCnt + '件');

  // 検証
  console.log('\n=== 検証 ===');
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

  // 5/2が無事か
  const nowExist = await drive.files.list({
    q: "name = '2026年5月2日' and trashed=false and mimeType='application/vnd.google-apps.spreadsheet'",
    fields: 'files(id,name)', pageSize: 5, includeItemsFromAllDrives: true, supportsAllDrives: true
  });
  console.log('\n5/2の日報:', nowExist.data.files.length, '件 (1件以上ならOK)');
  nowExist.data.files.forEach(f => console.log('  ' + f.name + ' | ' + f.id));
})().catch(e => { console.error('FATAL:', e.message); process.exit(1); });
