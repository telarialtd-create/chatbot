require('dotenv').config();
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const DOC_ID = '1tFfn9K2zNwxoeXxA2T_-aChmVEDqd69yemZh_k3P1z8';
const OUTPUT_PATH = path.join(__dirname, 'context/thoughts/product_vision.md');

function createAuthClient() {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    'http://localhost'
  );
  oauth2Client.setCredentials({
    access_token: process.env.GOOGLE_ACCESS_TOKEN,
    refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
  });
  return oauth2Client;
}

function extractText(content) {
  if (!content || !content.body || !content.body.content) return '';
  const lines = [];
  for (const elem of content.body.content) {
    if (!elem.paragraph) continue;
    const text = elem.paragraph.elements
      .map(e => (e.textRun ? e.textRun.content : ''))
      .join('');
    lines.push(text);
  }
  return lines.join('');
}

async function main() {
  const auth = createAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const res = await drive.files.export({
    fileId: DOC_ID,
    mimeType: 'text/plain',
  }, { responseType: 'text' });
  const text = res.data;

  const output = `# プロダクトビジョン・やりたいこと\n\n> Google Doc (${DOC_ID}) から自動取得\n> 最終更新: ${new Date().toLocaleString('ja-JP')}\n\n---\n\n${text}`;
  fs.writeFileSync(OUTPUT_PATH, output, 'utf8');
  console.log('保存完了:', OUTPUT_PATH);
  console.log('--- 内容プレビュー ---');
  console.log(text.slice(0, 500));
}

main().catch(err => {
  console.error('エラー:', err.message);
  process.exit(1);
});
