const { createAuthClient } = require('/app/chatbot/lib/common');
const { google } = require('googleapis');
const fs = require('fs');
(async () => {
  const auth = createAuthClient();
  const SCRIPT_ID = '1QCpsGYiQwFKKeVXw3GhTjZCFT4m1y3spVe8h15L0phpH9q7_gEpHbW_b';
  const script = google.script({ version: 'v1', auth });
  const r = await script.projects.getContent({ scriptId: SCRIPT_ID });
  const dir = '/app/chatbot/_gas_backup';
  fs.mkdirSync(dir, { recursive: true });
  const ts = new Date().toISOString().replace(/[:.]/g, '-');
  const meta = { savedAt: new Date().toISOString(), scriptId: SCRIPT_ID, files: r.data.files.map(f => ({ name: f.name, type: f.type })) };
  fs.writeFileSync(`${dir}/manifest_${ts}.json`, JSON.stringify(meta, null, 2));
  // 全文をjsonで保存
  fs.writeFileSync(`${dir}/full_${ts}.json`, JSON.stringify(r.data, null, 2));
  // 個別にも保存
  for (const f of r.data.files) {
    const ext = f.type === 'JSON' ? 'json' : 'gs';
    fs.writeFileSync(`${dir}/${f.name}_${ts}.${ext}`, f.source || '');
  }
  console.log('バックアップ保存:', dir);
  console.log('ファイル一覧:');
  r.data.files.forEach(f => console.log('  -', f.name + '.' + f.type, '(' + (f.source||'').length + '文字)'));

  // Code.gsを抽出してgrep的に「main」「trigger」「createTrigger」「14」「2週間」を表示
  const code = (r.data.files.find(f => f.name === 'Code') || {}).source || '';
  console.log('\n=== Code.gs内の関数名一覧 ===');
  const funcs = [];
  const reFunc = /^function\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(/gm;
  let m;
  while ((m = reFunc.exec(code))) funcs.push({ name: m[1], pos: m.index, line: code.slice(0, m.index).split('\n').length });
  funcs.forEach(f => console.log('  L' + f.line + ': function ' + f.name));

  console.log('\n=== "newTrigger"出現箇所 ===');
  const lines = code.split('\n');
  lines.forEach((l, i) => {
    if (/newTrigger|createTrigger|deleteTrigger|getProjectTriggers|TimeBased|everyDays|atHour|nearMinute/i.test(l)) {
      console.log('  L' + (i+1) + ': ' + l.trim().slice(0, 200));
    }
  });

  console.log('\n=== "14|2週間|Drive\\.copy|files\\.copy|makeCopy"出現箇所 ===');
  lines.forEach((l, i) => {
    if (/14|2週間|makeCopy|DriveApp.*copy|FilesResource|Drive\.Files\.copy/i.test(l)) {
      console.log('  L' + (i+1) + ': ' + l.trim().slice(0, 200));
    }
  });
})().catch(e => console.error('FATAL:', e.message, e.stack));
