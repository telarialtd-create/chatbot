const { createAuthClient } = require('/app/chatbot/lib/common');
const { google } = require('googleapis');
(async () => {
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  console.log('=== 単独検索: 2026年5月3日〜5月20日 ===');
  for (let d = 3; d <= 20; d++) {
    const name = '2026年5月' + d + '日';
    const res = await drive.files.list({
      q: "name = '" + name + "' and trashed=false",
      fields: 'files(id,name,createdTime,modifiedTime,parents,owners(emailAddress),shortcutDetails,mimeType)',
      pageSize: 20,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true
    });
    if (res.data.files.length > 0) {
      res.data.files.forEach(f => console.log(f.name + ' | id=' + f.id + ' | parent=' + (f.parents||[]).join(',') + ' | created=' + f.createdTime + ' | owner=' + (f.owners||[]).map(o=>o.emailAddress).join(',') + ' | mime=' + f.mimeType + (f.shortcutDetails ? ' | SHORTCUT->' + f.shortcutDetails.targetId : '')));
    }
  }
  console.log('\n=== 日報フォルダ16R1BK5Nnv... 全ファイル ===');
  const all = await drive.files.list({
    q: "'16R1BK5NnvYkH4Eqh6t51OGQl0tXVJ3If' in parents and trashed=false",
    fields: 'files(id,name,createdTime,modifiedTime,owners(emailAddress),mimeType,shortcutDetails)',
    pageSize: 100,
    orderBy: 'name',
    includeItemsFromAllDrives: true,
    supportsAllDrives: true
  });
  console.log('合計:', all.data.files.length, '件');
  all.data.files.forEach(f => console.log('  ' + f.name + ' | created=' + f.createdTime + ' | owner=' + (f.owners||[]).map(o=>o.emailAddress).join(',') + ' | mime=' + f.mimeType + (f.shortcutDetails ? ' | SHORTCUT->' + f.shortcutDetails.targetId : '')));
})().catch(e => { console.error('ERR:', e.message); });
