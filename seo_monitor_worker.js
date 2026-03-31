/**
 * seo_monitor_worker.js
 * 求人キーワードでGoogle検索 → 1ページ目の各サイトにアクセス
 * CREA・ふわもこSPAが掲載されていない場合にLINEで通知
 *
 * チェック間隔: 6時間ごと
 */

require('dotenv').config();
const puppeteer = require('puppeteer');
const { messagingApi } = require('@line/bot-sdk');

// ── 設定 ──────────────────────────────────────────────────
const KEYWORDS = [
  'メンズエステ求人',
  'メンズエステ　求人',
  'メンエス求人',
  'メンエス　求人',
];

// 自店舗名マッチパターン
const STORE_PATTERNS = [
  /CREA/i,
  /クレア/,
  /ふわもこ/,
  /fuwamoko/i,
];

const CHECK_INTERVAL_MS = 6 * 60 * 60 * 1000; // 6時間

// ── LINE クライアント ─────────────────────────────────────
function createLineClient() {
  return new messagingApi.MessagingApiClient({
    channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  });
}

// ── 文字列に自店舗名が含まれるか判定 ─────────────────────
function containsStore(text) {
  return STORE_PATTERNS.some(pattern => pattern.test(text));
}

// ── Google検索で1ページ目のURLリストを取得 ───────────────
async function getFirstPageUrls(browser, keyword) {
  const page = await browser.newPage();
  await page.setUserAgent(
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
  );

  try {
    const searchUrl = `https://www.google.co.jp/search?q=${encodeURIComponent(keyword)}&num=10&hl=ja`;
    await page.goto(searchUrl, { waitUntil: 'networkidle2', timeout: 30000 });

    // 検索結果のリンク・タイトル・スニペットを取得
    const results = await page.evaluate(() => {
      const items = [];
      const elements = document.querySelectorAll('#search .g, #rso > div');
      elements.forEach(el => {
        const titleEl   = el.querySelector('h3');
        const snippetEl = el.querySelector('.VwiC3b, [data-sncf="1"], .s3v9rd');
        const linkEl    = el.querySelector('a[href^="http"]');
        if (!linkEl) return;
        items.push({
          title:   titleEl?.textContent?.trim()   || '',
          snippet: snippetEl?.textContent?.trim() || '',
          url:     linkEl.href,
        });
      });
      return items;
    });

    console.log(`[SEO] "${keyword}" → ${results.length}件取得`);
    return results;
  } finally {
    await page.close();
  }
}

// ── 各サイトにアクセスして自店舗掲載を確認 ───────────────
async function checkSiteForStore(browser, url) {
  const page = await browser.newPage();
  await page.setUserAgent(
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
  );

  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 20000 });
    const bodyText = await page.evaluate(() => document.body?.innerText || '');
    return containsStore(bodyText);
  } catch (err) {
    console.warn(`[SEO] サイトアクセス失敗 (${url.slice(0, 60)}...): ${err.message}`);
    return false;
  } finally {
    await page.close();
  }
}

// ── 1キーワード分のチェック ──────────────────────────────
async function checkKeyword(browser, keyword) {
  // Step1: Google検索結果を取得
  const results = await getFirstPageUrls(browser, keyword);
  if (results.length === 0) {
    console.warn(`[SEO] "${keyword}" 検索結果が0件（Google側ブロックの可能性）`);
    return { keyword, found: false, foundIn: null, reason: '検索結果0件' };
  }

  // Step2: 検索結果のタイトル・スニペットで即チェック
  for (const r of results) {
    const text = `${r.title} ${r.snippet} ${r.url}`;
    if (containsStore(text)) {
      console.log(`[SEO] "${keyword}" → 検索結果に掲載あり: ${r.title}`);
      return { keyword, found: true, foundIn: r.url };
    }
  }

  // Step3: 各サイトを実際に訪問してチェック
  console.log(`[SEO] "${keyword}" → 検索結果スニペットに見当たらず、各サイトを訪問します`);
  for (const r of results) {
    // Google自身・広告系は除外
    if (r.url.includes('google.') || r.url.includes('doubleclick.')) continue;

    console.log(`[SEO]   訪問: ${r.url.slice(0, 70)}`);
    const found = await checkSiteForStore(browser, r.url);
    if (found) {
      console.log(`[SEO] "${keyword}" → "${r.title}" に掲載あり`);
      return { keyword, found: true, foundIn: r.url };
    }
    // サイト間のアクセス間隔
    await new Promise(r => setTimeout(r, 2000));
  }

  console.log(`[SEO] "${keyword}" → 1ページ目のどのサイトにも掲載なし`);
  return { keyword, found: false, foundIn: null };
}

// ── LINEで通知 ────────────────────────────────────────────
async function sendLineNotification(missingKeywords) {
  const client = createLineClient();
  const target = process.env.LINE_USER_ID;

  const jstNow = new Date(Date.now() + 9 * 60 * 60 * 1000);
  const timeStr = jstNow.toISOString().replace('T', ' ').slice(0, 16);

  const msg = [
    '【求人SEO 掲載なし通知】',
    '',
    '以下のキーワードでCREA・ふわもこSPAが',
    'Google 1ページ目に見当たりませんでした:',
    '',
    ...missingKeywords.map(k => `・${k}`),
    '',
    `確認日時(JST): ${timeStr}`,
  ].join('\n');

  await client.pushMessage({
    to: target,
    messages: [{ type: 'text', text: msg }],
  });
  console.log('[SEO] LINE通知送信完了');
}

// ── メインチェック処理 ────────────────────────────────────
async function checkAndNotify() {
  const jstNow = new Date(Date.now() + 9 * 60 * 60 * 1000);
  console.log(`[SEO] ===== チェック開始: ${jstNow.toISOString().slice(0, 16)} =====`);

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  const missingKeywords = [];

  try {
    for (const keyword of KEYWORDS) {
      const result = await checkKeyword(browser, keyword);
      if (!result.found) {
        missingKeywords.push(keyword);
      }
      // キーワード間のインターバル（Google対策）
      await new Promise(r => setTimeout(r, 5000));
    }
  } finally {
    await browser.close();
  }

  if (missingKeywords.length > 0) {
    await sendLineNotification(missingKeywords);
  } else {
    console.log('[SEO] 全キーワードで掲載確認済み。通知なし。');
  }
}

// ── メインループ ─────────────────────────────────────────
async function main() {
  console.log('[SEO] 求人SEOモニター起動');
  console.log('[SEO] 監視キーワード:', KEYWORDS);
  console.log(`[SEO] チェック間隔: ${CHECK_INTERVAL_MS / 3600000}時間`);

  while (true) {
    try {
      await checkAndNotify();
    } catch (err) {
      console.error('[SEO] エラー:', err.message);
    }

    console.log(`[SEO] 次回チェックまで ${CHECK_INTERVAL_MS / 3600000} 時間待機`);
    await new Promise(r => setTimeout(r, CHECK_INTERVAL_MS));
  }
}

main().catch(err => {
  console.error('[SEO] 致命的エラー:', err);
  process.exit(1);
});
