/**
 * realtime_worker.js
 * エステ魂 リアルタイム自動更新
 *
 * サイクル:
 *   ご案内状況       → 1分10秒ごと（毎ループ）
 *   リアルタイム集客  → 10分ごと
 *   リアルタイム求人  → 10分ごと
 *   停止時間         → 5:00〜10:00
 */

require('dotenv').config();
const puppeteer = require('puppeteer');

const ANNNAI_INTERVAL_MS   = 70 * 1000;       // 1分10秒
const REALTIME_INTERVAL_MS = (10 * 60 + 30) * 1000;  // 10分30秒
const PAUSE_START = { hour: 6, minute: 5 };
const PAUSE_END   = { hour: 7, minute: 10 };

// テキストを含む要素を探してクリック
async function clickByText(page, text) {
  const clicked = await page.evaluate((text) => {
    const tags = ['a', 'button', 'li', 'div', 'span', 'input[type="submit"]'];
    for (const tag of tags) {
      const els = [...document.querySelectorAll(tag)];
      const el = els.find(e => e.innerText && e.innerText.trim().includes(text));
      if (el) { el.click(); return true; }
    }
    return false;
  }, text);

  if (!clicked) throw new Error(`要素が見つかりません: "${text}"`);
  await new Promise(r => setTimeout(r, 1500));
}

// ページ遷移を待つ（遷移しない場合もある）
async function safeWaitForNav(page, timeout = 8000) {
  try {
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout });
  } catch (_) {}
}

// ログイン
async function login(page) {
  const user = process.env.ESTAMA_USER || '';
  const pass = process.env.ESTAMA_PASS || '';
  console.log('[login] ログイン中... user=' + (user ? user.substring(0, 5) + '***' : 'EMPTY'));
  await page.goto('https://estama.jp/login/?r=/admin/', { waitUntil: 'networkidle2' });
  await page.type('#inputEmail', user, { delay: 30 });
  await page.type('#inputPassword', pass, { delay: 30 });
  await Promise.all([
    page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }),
    page.click('a[data-post="login_shop"]'),
  ]);
  console.log('[login] ログイン完了');
}

// セッション切れ確認（ログインページにいたら再ログイン）
async function ensureLoggedIn(page) {
  if (page.url().includes('/login/')) {
    console.log('[auth] セッション切れ → 再ログイン');
    await login(page);
  }
}

// ─────────────────────────────────────────
// ① ご案内状況 → ブースト → 閉じる
// ─────────────────────────────────────────
async function runAnnaijokyo(page) {
  console.log('[案内状況] 開始');
  await page.goto('https://estama.jp/admin/', { waitUntil: 'networkidle2' });
  await ensureLoggedIn(page);

  await clickByText(page, 'ご案内状況');
  await safeWaitForNav(page);

  await clickByText(page, '今すぐご案内可');
  await new Promise(r => setTimeout(r, 1500));

  await clickByText(page, '閉じる');
  await new Promise(r => setTimeout(r, 1000));

  console.log('[案内状況] 完了');
}

// ─────────────────────────────────────────
// ② リアルタイム集客 → 新規投稿 → テンプレ選択 → 投稿
// ─────────────────────────────────────────
async function runRealtimeKyakka(page) {
  console.log('[集客] 開始');
  await page.goto('https://estama.jp/admin/', { waitUntil: 'networkidle2' });
  await ensureLoggedIn(page);

  await clickByText(page, 'リアルタイム集客');
  await safeWaitForNav(page);

  await clickByText(page, '新しいメッセージを書く');
  await safeWaitForNav(page);

  // テンプレート選択
  const selected = await page.evaluate(() => {
    const option = [...document.querySelectorAll('option')]
      .find(e => e.textContent.includes('新人入店記念キャンペーン'));
    if (option) {
      option.selected = true;
      option.closest('select').dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    }
    return false;
  });
  if (!selected) console.warn('[集客] テンプレートが見つかりませんでした');
  await new Promise(r => setTimeout(r, 1500));

  await clickByText(page, '投稿する');
  await safeWaitForNav(page);

  console.log('[集客] 完了');
}

// ─────────────────────────────────────────
// ③ リアルタイム求人 → 複製 → 投稿
// ─────────────────────────────────────────
async function runRealtimeKyujin(page) {
  console.log('[求人] 開始');
  await page.goto('https://estama.jp/admin/', { waitUntil: 'networkidle2' });
  await ensureLoggedIn(page);

  await clickByText(page, 'リアルタイム求人');
  await safeWaitForNav(page);

  await clickByText(page, '複製');
  await safeWaitForNav(page);

  await clickByText(page, '投稿する');
  await safeWaitForNav(page);

  console.log('[求人] 完了');
}

// ─────────────────────────────────────────
// メインループ
// ─────────────────────────────────────────
async function main() {
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 900 });
  await page.setUserAgent('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36');

  await login(page);

  let lastRealtimeRun = 0;

  while (true) {
    const now = new Date();

    // 6:05〜7:10 は停止
    const nowMins = now.getHours() * 60 + now.getMinutes();
    const pauseStart = PAUSE_START.hour * 60 + PAUSE_START.minute;
    const pauseEnd   = PAUSE_END.hour   * 60 + PAUSE_END.minute;
    if (nowMins >= pauseStart && nowMins < pauseEnd) {
      console.log(`[停止中] ${now.toTimeString().slice(0,5)} (6:05〜7:10)`);
      await new Promise(r => setTimeout(r, 60 * 1000));
      continue;
    }

    try {
      // 10分ごと: リアルタイム集客 + 求人
      if (Date.now() - lastRealtimeRun >= REALTIME_INTERVAL_MS) {
        await runRealtimeKyakka(page);
        await runRealtimeKyujin(page);
        lastRealtimeRun = Date.now();
      }

      // 毎ループ: ご案内状況
      await runAnnaijokyo(page);

    } catch (err) {
      console.error('[エラー]', err.message);
      // エラー時は再ログインを試みる
      try {
        await login(page);
      } catch (e) {
        console.error('[再ログイン失敗]', e.message);
      }
    }

    const jitter = Math.floor(Math.random() * 20 - 10) * 1000; // ±10秒のゆらぎ
    const wait = ANNNAI_INTERVAL_MS + jitter;
    console.log(`[待機] ${Math.round(wait / 1000)}秒後に再実行`);
    await new Promise(r => setTimeout(r, wait));
  }
}

main().catch(err => {
  console.error('[致命的エラー]', err);
  process.exit(1);
});
