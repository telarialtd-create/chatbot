/**
 * store_worker.js
 * 1店舗専用ワーカー
 * 環境変数 STORE_* から店舗設定を読み込んで動作する
 *
 * 機能:
 *   - 案内状況更新（毎分）
 *   - リアルタイム集客（10分ごと）
 *   - リアルタイム求人（10分ごと）
 *   - 写メ日記自動投稿（10回/日 JST 10:30〜01:57）
 *   - 店長ブログ自動投稿（10回/日 同上）
 */

require('dotenv').config();
const puppeteer = require('puppeteer');
const { google } = require('googleapis');
const fs = require('fs');

// ─────────────────────────────────────────
// 設定（既存）
// ─────────────────────────────────────────
const ANNNAI_INTERVAL_MS   = (parseInt(process.env.STORE_ANNNAI_MIN   || '1',  10) * 60 + parseInt(process.env.STORE_ANNNAI_SEC   || '0', 10)) * 1000;
const REALTIME_INTERVAL_MS = (parseInt(process.env.STORE_REALTIME_MIN || '10', 10) * 60 + parseInt(process.env.STORE_REALTIME_SEC || '0', 10)) * 1000;
const KYUJIN_INTERVAL_MS   = (parseInt(process.env.STORE_KYUJIN_MIN   || '10', 10) * 60 + parseInt(process.env.STORE_KYUJIN_SEC   || '0', 10)) * 1000;
const STORE_START_TIME = process.env.STORE_START_TIME || '07:10';
const STORE_END_TIME   = process.env.STORE_END_TIME   || '06:05';

// ─────────────────────────────────────────
// 設定（写メ日記・店長ブログ）
// ─────────────────────────────────────────
const SHINSHA_ENABLED  = process.env.STORE_SHINSHA_ENABLED === 'true';
const BLOG_ENABLED     = process.env.STORE_BLOG_ENABLED === 'true';
const TAIKEN_ENABLED   = process.env.STORE_TAIKEN_ENABLED === 'true';
const SHOP_ID          = process.env.STORE_SHOP_ID || '';
const RESERVE_URL      = process.env.STORE_RESERVE_URL || '';
const SHINSHA_TITLE    = process.env.STORE_SHINSHA_TITLE || '★CREAの出勤情報★';
const BLOG_TITLE       = process.env.STORE_BLOG_TITLE   || '★CREA広島の求人情報★';
const BLOG_CONTENT     = process.env.STORE_BLOG_CONTENT || '';

// 画像フォルダID（シートのAA列から動的に読み込み、env varにフォールバック）
let _imageFolderId = null;
async function getImageFolderId() {
  if (_imageFolderId !== null) return _imageFolderId;
  try {
    const sheets    = await getSheetsClient();
    const row       = await findStoreRow();
    if (row) {
      const sheetName = process.env.STORES_SHEET_NAME || '店舗設定';
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.STORES_SPREADSHEET_ID,
        range: `'${sheetName}'!AA${row}`,
      });
      const val = ((res.data.values || [[]])[0] || [])[0] || '';
      if (val) {
        _imageFolderId = val;
        log(`画像フォルダID（シートから取得）: ${val}`);
        return _imageFolderId;
      }
    }
  } catch (_) {}
  _imageFolderId = process.env.STORE_IMAGE_FOLDER_ID || '';
  if (_imageFolderId) log(`画像フォルダID（env変数）: ${_imageFolderId}`);
  return _imageFolderId;
}

// 投稿スケジュール（JST 10:30〜01:57 均等10回）
const POST_SCHEDULE_MINS = [630, 733, 836, 939, 1042, 1145, 1248, 1351, 14, 117];
//                          10:30 12:13 13:56 15:39 17:22 19:05 20:48 22:31 00:14 01:57
const POST_WINDOW_MINS = 3; // ±3分以内で実行

let postedSlotsToday  = new Set();
let lastPostedDateKey = null;
let taikenDoneDate    = null;

// ─────────────────────────────────────────
// 時刻ユーティリティ
// ─────────────────────────────────────────
function parseTimeMins(timeStr) {
  const [h, m] = timeStr.split(':').map(Number);
  return (h || 0) * 60 + (m || 0);
}

function isRunningTime(now) {
  const jstMins   = (now.getUTCHours() * 60 + now.getUTCMinutes() + 9 * 60) % (24 * 60);
  const nowMins   = jstMins;
  const startMins = parseTimeMins(STORE_START_TIME);
  const endMins   = parseTimeMins(STORE_END_TIME);
  if (startMins <= endMins) {
    return nowMins >= startMins && nowMins < endMins;
  } else {
    return nowMins >= startMins || nowMins < endMins;
  }
}

function getJSTInfo() {
  const now    = new Date();
  const jstMs  = now.getTime() + 9 * 60 * 60 * 1000;
  const jst    = new Date(jstMs);
  const h      = jst.getUTCHours();
  const m      = jst.getUTCMinutes();
  const totalMins = h * 60 + m;
  // 6時前は前日扱い（1日の区切りを06:00とする）
  const baseDate = h < 6 ? new Date(jstMs - 24 * 60 * 60 * 1000) : jst;
  const dateKey  = `${baseDate.getUTCFullYear()}-${baseDate.getUTCMonth() + 1}-${baseDate.getUTCDate()}`;
  return { totalMins, dateKey };
}

function getDuePostSlotIndex() {
  const { totalMins, dateKey } = getJSTInfo();
  // 日付が変わったらリセット
  if (lastPostedDateKey !== dateKey) {
    postedSlotsToday  = new Set();
    lastPostedDateKey = dateKey;
  }
  for (let i = 0; i < POST_SCHEDULE_MINS.length; i++) {
    if (postedSlotsToday.has(i)) continue;
    const slotMins = POST_SCHEDULE_MINS[i];
    let diff = Math.abs(totalMins - slotMins);
    diff = Math.min(diff, 24 * 60 - diff); // 深夜またぎ対応
    if (diff <= POST_WINDOW_MINS) return i;
  }
  return -1;
}

function nowJST() {
  return new Date().toLocaleTimeString('ja-JP', { timeZone: 'Asia/Tokyo', hour: '2-digit', minute: '2-digit' });
}

// ─────────────────────────────────────────
// 店舗設定
// ─────────────────────────────────────────
const store = {
  name:            process.env.STORE_NAME             || '',
  user:            process.env.STORE_USER             || '',
  pass:            process.env.STORE_PASS             || '',
  proxyHost:       process.env.STORE_PROXY_HOST       || '',
  proxyPort:       process.env.STORE_PROXY_PORT       || '',
  proxyUser:       process.env.STORE_PROXY_USER       || '',
  proxyPass:       process.env.STORE_PROXY_PASS       || '',
  templateKeyword: process.env.STORE_TEMPLATE_KEYWORD || '',
};

if (!store.user || !store.pass) {
  console.error('[エラー] STORE_USER / STORE_PASS が未設定です');
  process.exit(1);
}

function log(msg) {
  const time = new Date().toLocaleTimeString('ja-JP', { timeZone: 'Asia/Tokyo' });
  console.log(`[${time}][${store.name}] ${msg}`);
}

// ─────────────────────────────────────────
// Google Sheets ステータス更新
// ─────────────────────────────────────────
let _sheetsClient = null;
let _storeRow     = null;

async function getSheetsClient() {
  if (_sheetsClient) return _sheetsClient;
  const auth = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
  );
  const refreshToken = process.env.GOOGLE_WRITE_REFRESH_TOKEN || process.env.GOOGLE_REFRESH_TOKEN;
  auth.setCredentials({ refresh_token: refreshToken });
  _sheetsClient = google.sheets({ version: 'v4', auth });
  return _sheetsClient;
}

async function findStoreRow() {
  if (_storeRow) return _storeRow;
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.STORES_SPREADSHEET_ID,
    range: `'${process.env.STORES_SHEET_NAME || '店舗設定'}'!A2:A200`,
  });
  const rows = res.data.values || [];
  const idx  = rows.findIndex(r => r[0] === store.name);
  if (idx >= 0) _storeRow = idx + 2;
  return _storeRow;
}

async function updateSheetStatus(col, text) {
  if (!process.env.STORES_SPREADSHEET_ID || !process.env.GOOGLE_REFRESH_TOKEN) return;
  try {
    const sheets    = await getSheetsClient();
    const row       = await findStoreRow();
    if (!row) return;
    const sheetName = process.env.STORES_SHEET_NAME || '店舗設定';
    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.STORES_SPREADSHEET_ID,
      range: `'${sheetName}'!${col}${row}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[text]] },
    });
  } catch (_) {}
}

// ─────────────────────────────────────────
// Google Drive 画像ダウンロード
// ─────────────────────────────────────────
async function downloadRandomImage(folderId) {
  if (!folderId) return null;
  try {
    const auth = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
    );
    auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
    const drive = google.drive({ version: 'v3', auth });

    const res = await drive.files.list({
      q: `'${folderId}' in parents and mimeType contains 'image/' and trashed=false`,
      fields: 'files(id,name)',
    });
    const files = res.data.files;
    if (!files || files.length === 0) {
      log('画像フォルダが空です（画像を追加してください）');
      return null;
    }

    const file    = files[Math.floor(Math.random() * files.length)];
    const ext     = (file.name.split('.').pop() || 'jpg').toLowerCase();
    const tmpPath = `/tmp/estama_img_${Date.now()}.${ext}`;

    const dlRes = await drive.files.get(
      { fileId: file.id, alt: 'media' },
      { responseType: 'stream' },
    );
    const dest = fs.createWriteStream(tmpPath);
    await new Promise((resolve, reject) => {
      dlRes.data.pipe(dest).on('finish', resolve).on('error', reject);
    });

    log(`画像ダウンロード完了: ${file.name}`);
    return tmpPath;
  } catch (err) {
    log(`画像ダウンロードエラー: ${err.message}`);
    return null;
  }
}

// ─────────────────────────────────────────
// 出勤スケジュールのスクレイピング
// ─────────────────────────────────────────
async function fetchTodaySchedule(browser, shopId) {
  const p = await browser.newPage();
  try {
    // 管理画面の出勤表（1日表示）から全スタッフ取得
    await p.goto('https://estama.jp/admin/schedule/', {
      waitUntil: 'networkidle2',
      timeout: 20000,
    });
    await ensureLoggedIn(p);

    const results = await p.evaluate(() => {
      const lines   = document.body.innerText.split('\n').map(l => l.trim()).filter(l => l);
      const timeRe  = /^\d{1,2}:\d{2}$/;
      const statusRe = /^[×○─]$|^TEL$/;
      const skipRe  = /前|次|週|全|設定|出勤|受付|表示|切替|印刷|追加|お気に入り|コピー|ページ|estama|管理|集客|求人|アクセス|シェア|営業|セラピスト|ご案内|確認|時間|Copyright/;
      const dayOfWeekRe = /^[（(][月火水木金土日][）)]$|^[月火水木金土日]$/;

      // 最初の時刻行を検出
      const firstTimeIdx = lines.findIndex(l => timeRe.test(l));
      if (firstTimeIdx < 0) return [];

      // 時刻行直前からスタッフ名を抽出
      const staffNames = [];
      const seen = new Set();
      for (let i = Math.max(0, firstTimeIdx - 40); i < firstTimeIdx; i++) {
        const line = lines[i];
        if (
          line.length >= 2 && line.length <= 15 &&
          !skipRe.test(line) && !dayOfWeekRe.test(line) && !seen.has(line) &&
          !/^\d/.test(line) && !/[a-zA-Z]/.test(line)
        ) {
          seen.add(line);
          staffNames.push(line);
        }
      }

      // 時刻ブロックを解析して各スタッフの勤務時間を算出
      const staffSchedule = {};
      let i = firstTimeIdx;
      while (i < lines.length) {
        if (timeRe.test(lines[i])) {
          const timeStr = lines[i];
          const statuses = [];
          let j = i + 1;
          while (j < lines.length && statuses.length < staffNames.length) {
            if (statusRe.test(lines[j]))  statuses.push(lines[j]);
            else if (timeRe.test(lines[j])) break;
            j++;
          }
          staffNames.forEach((name, idx) => {
            const status = statuses[idx] || '─';
            if (status !== '─') {
              if (!staffSchedule[name]) staffSchedule[name] = { first: timeStr, last: timeStr };
              else staffSchedule[name].last = timeStr;
            }
          });
          i = j;
        } else {
          i++;
        }
      }

      // 勤務開始時刻順にソートして返す
      return staffNames
        .filter(name => staffSchedule[name])
        .map(name => ({ name, time: `${staffSchedule[name].first}～${staffSchedule[name].last}` }))
        .sort((a, b) => {
          const toMins = t => { const [h, m] = t.split(':').map(Number); return h * 60 + m; };
          return toMins(a.time.split('～')[0]) - toMins(b.time.split('～')[0]);
        });
    });

    log(`スケジュール取得（管理画面）: ${results.length}名`);
    return results;
  } catch (err) {
    log(`スケジュール取得エラー: ${err.message}`);
    return [];
  } finally {
    await p.close();
  }
}

function buildShinshaBody(staffList) {
  const lines = ['【本日の出勤情報】', ''];
  if (staffList.length > 0) {
    staffList.forEach(s => lines.push(`★${s.name}　${s.time}`));
  } else {
    lines.push('本日の出勤情報は準備中です');
  }
  lines.push('');
  lines.push('●ご予約はこちら');
  lines.push(RESERVE_URL || `https://estama.jp/shop/${SHOP_ID}/reserve/`);
  lines.push('');
  lines.push('#メンズエステ #エステ魂');
  return lines.join('\n');
}

// ─────────────────────────────────────────
// Puppeteer ヘルパー
// ─────────────────────────────────────────
async function clickByText(page, text) {
  const clicked = await page.evaluate((text) => {
    const tags = ['a', 'button', 'li', 'div', 'span', 'input[type="submit"]'];
    for (const tag of tags) {
      const els = [...document.querySelectorAll(tag)];
      const el  = els.find(e => e.innerText && e.innerText.trim().includes(text));
      if (el) { el.click(); return true; }
    }
    return false;
  }, text);
  if (!clicked) throw new Error(`要素が見つかりません: "${text}"`);
  await new Promise(r => setTimeout(r, 1500));
}

async function safeWaitForNav(page, timeout = 8000) {
  try {
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout });
  } catch (_) {}
}

async function login(page) {
  log(`ログイン中... user=${store.user.substring(0, 5)}***`);
  await page.goto('https://estama.jp/login/?r=/admin/', { waitUntil: 'networkidle2' });
  await page.type('#inputEmail', store.user, { delay: 30 });
  await page.type('#inputPassword', store.pass, { delay: 30 });
  await Promise.all([
    page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }),
    page.click('a[data-post="login_shop"]'),
  ]);
  log('ログイン完了');
}

async function ensureLoggedIn(page) {
  if (page.url().includes('/login/')) {
    log('セッション切れ → 再ログイン');
    await login(page);
  }
}

// ─────────────────────────────────────────
// 画像アップロード（fileInput / croppic方式）フォールバック
// ─────────────────────────────────────────
async function uploadImageByFileInput(page, imgPath) {
  try {
    // hidden fields をセット（croppicがCSRFを自動付与）
    await page.evaluate(() => {
      const $ = window.jQuery;
      if ($) {
        $('#img-crop-id').val('blog_icon_1');
        $('#img-crop-target').val('blog_icon_1-imgupload');
        $('#img-crop-mode').val('one');
      }
    });
    const fileInput = await page.$('#blog_icon_1-imgupload_imgUploadField');
    if (!fileInput) { log('fileInput見つからず'); return; }
    await fileInput.uploadFile(imgPath);
    log('fileInput方式: アップロード中...');
    // モーダル or croppedImgを待機（どちらかが出れば成功）
    await Promise.race([
      page.waitForSelector('.cropControlCrop', { timeout: 30000 }),
      page.waitForSelector('#croppicModal .cropControlCrop', { timeout: 30000 }),
    ]);
    await new Promise(r => setTimeout(r, 500));
    await page.evaluate(() => {
      const btn = document.querySelector('#croppicModal .cropControlCrop')
               || document.querySelector('.cropControlCrop');
      if (btn && window.jQuery) window.jQuery(btn).trigger('click');
      else if (btn) btn.click();
    });
    await page.waitForSelector('#blog_icon_1-imgupload .croppedImg', { timeout: 15000 });
    log('fileInput方式: クロップ完了');
  } catch (err) {
    log(`fileInput方式 エラー: ${err.message}`);
  }
}

// ─────────────────────────────────────────
// 画像アップロードヘルパー（XHR直接POST方式）
// ─────────────────────────────────────────
async function uploadImageToForm(page, imgPath) {
  if (!imgPath) return;
  try {
    log('画像アップロード中（XHR直接POST方式）...');

    // Node.js でファイルをbase64に変換してブラウザに渡す
    const fileBuffer = fs.readFileSync(imgPath);
    const base64Data = fileBuffer.toString('base64');
    const ext = (imgPath.split('.').pop() || 'jpg').toLowerCase();
    const mimeType = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : 'image/jpeg';
    const fileName = require('path').basename(imgPath);

    const currentUrl = page.url();
    const origin = new URL(currentUrl).origin;

    // ブラウザ内XHRで /post/uptemp/ に直接POST（CSRF・Cookieを自動付与）
    const uptempResult = await page.evaluate(async (base64, mime, fname, origin, currentUrl) => {
      return new Promise((resolve) => {
        const byteStr = atob(base64);
        const ab = new ArrayBuffer(byteStr.length);
        const ia = new Uint8Array(ab);
        for (let i = 0; i < byteStr.length; i++) ia[i] = byteStr.charCodeAt(i);
        const blob = new Blob([ab], { type: mime });

        // estama.jpはCSRFトークンとして ctk_cookie（cookie）と ctk（hidden input）を使用
        const ctkCookie = document.cookie.split(';')
          .map(c => c.trim().split('='))
          .find(([k]) => k && k.trim() === 'ctk_cookie');
        const ctkInput = document.querySelector('input[name="ctk"]');
        const ctkValue = (ctkCookie ? decodeURIComponent(ctkCookie[1] || '') : null)
                      || (ctkInput ? ctkInput.value : null);

        const form = new FormData();
        form.append('img', blob, fname);  // croppicのフィールド名は 'img'
        form.append('id', 'blog_icon_1');
        form.append('mode', 'one');
        form.append('target', 'blog_icon_1-imgupload');
        if (ctkValue) {
          form.append('ctk', ctkValue);
        }

        const xhr = new XMLHttpRequest();
        xhr.open('POST', `${origin}/post/uptemp/`, true);
        xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
        xhr.withCredentials = true;
        xhr.timeout = 25000;
        xhr.onload = () => {
          try { resolve(JSON.parse(xhr.responseText)); }
          catch (_) { resolve({ status: 'error', raw: xhr.responseText.substring(0, 200) }); }
        };
        xhr.onerror = () => resolve({ status: 'xhr_error', message: 'network error' });
        xhr.ontimeout = () => resolve({ status: 'xhr_timeout' });
        xhr.send(form);
      });
    }, base64Data, mimeType, fileName, origin, currentUrl);

    log(`[uptemp] ${JSON.stringify(uptempResult).substring(0, 150)}`);

    if (!uptempResult || uptempResult.status !== 'success' || !uptempResult.url) {
      log(`uptempアップロード失敗: ${JSON.stringify(uptempResult).substring(0, 150)}`);
      return;
    }

    const uptempUrl = uptempResult.url;
    log(`一時URL取得: ${uptempUrl}`);

    // クロップ: ブラウザfetchで直接POST（XHR方式ではcroppic UIが開かないため直接）
    const cropResult = await page.evaluate(async (origin, currentUrl, tempUrl) => {
      // estama.jpはCSRFトークンとして ctk_cookie（cookie）と ctk（hidden input）を使用
      const ctkCookie = document.cookie.split(';')
        .map(c => c.trim().split('='))
        .find(([k]) => k && k.trim() === 'ctk_cookie');
      const ctkInput = document.querySelector('input[name="ctk"]');
      const ctkValue = (ctkCookie ? decodeURIComponent(ctkCookie[1] || '') : null)
                    || (ctkInput ? ctkInput.value : null);

      const form = new FormData();
      ['imgInitW', 'imgInitH', 'imgW', 'imgH'].forEach(k => form.append(k, '400'));
      ['imgY1', 'imgX1'].forEach(k => form.append(k, '0'));
      ['cropW', 'cropH'].forEach(k => form.append(k, '400'));
      form.append('rotation', '0');
      form.append('imgUrl', tempUrl);   // croppicは 'imgUrl' フィールドを使用
      form.append('upload_id', 'blog_icon_1');
      form.append('id', 'blog_icon_1');
      form.append('mode', 'one');
      form.append('target', 'blog_icon_1-imgupload');
      if (ctkValue) form.append('ctk', ctkValue);

      try {
        const resp = await fetch(`${origin}/post/cropping/`, {
          method: 'POST',
          credentials: 'include',
          headers: { 'X-Requested-With': 'XMLHttpRequest', 'Referer': currentUrl },
          body: form,
        });
        const text = await resp.text();
        try { return JSON.parse(text); } catch (_) { return { status: 'error', raw: text.substring(0, 200) }; }
      } catch (e) {
        return { status: 'error', message: e.message };
      }
    }, origin, currentUrl, uptempUrl);

    if (!cropResult || cropResult.status !== 'success' || !cropResult.url) {
      log(`クロップ失敗: ${JSON.stringify(cropResult).substring(0, 200)}`);
      return;
    }

    const croppedUrl = cropResult.url;
    log(`クロップ完了: ${croppedUrl}`);

    // DOM に croppedImg と hidden input をセット
    // my_crop.jsの onAfterImgCrop と同じ処理：
    //   name="blog_icon_1-imgupload" の hidden input を .img_area 内に追加
    //   croppedImg を blog_icon_1-imgupload div に追加（表示用）
    await page.evaluate((url) => {
      // ① croppic div に croppedImg を追加（UI表示用）
      const c = document.getElementById('blog_icon_1-imgupload');
      if (c) {
        // 既存のcroppedImgを削除してから追加
        const existImg = c.querySelector('.croppedImg');
        if (existImg) existImg.remove();
        const img = document.createElement('img');
        img.className = 'croppedImg';
        img.src = url;
        c.appendChild(img);
      }

      // ② .img_area 内の up_items に hidden input を追加（フォーム送信用）
      // addoneimg と同等の処理: blog_icon_1 の親 div(.img_area) に up_items を prepend
      const blogIcon = document.getElementById('blog_icon_1');
      if (blogIcon) {
        const imgArea = blogIcon.closest('.img_area') || blogIcon.parentElement;
        // 既存の up_items を削除
        const existItems = imgArea.querySelector('.up_items');
        if (existItems) existItems.remove();
        // new up_items を作成
        const div = document.createElement('div');
        div.className = 'up_items';
        div.innerHTML =
          '<img src="' + url + '">' +
          '<a href="javascript:void(0)" class="temp-delete">× キャンセル</a>' +
          '<input name="blog_icon_1-imgupload" value="' + url + '" type="hidden">';
        imgArea.prepend(div);
        // img_upped クラスを追加
        imgArea.closest('.up_img_one')?.classList.add('img_upped');
      }
    }, croppedUrl);

    log('画像アップロード完了');
  } catch (err) {
    log(`画像アップロードエラー: ${err.message}`);
  } finally {
    try { fs.unlinkSync(imgPath); } catch (_) {}
  }
}

// ─────────────────────────────────────────
// 案内状況更新
// ─────────────────────────────────────────
async function runAnnaijokyo(page) {
  try {
    await page.goto('https://estama.jp/admin/', { waitUntil: 'networkidle2' });
    await ensureLoggedIn(page);
    await clickByText(page, 'ご案内状況');
    await safeWaitForNav(page);
    await clickByText(page, '今すぐご案内可');
    await new Promise(r => setTimeout(r, 1500));
    await page.evaluate(() => {
      const btn = [...document.querySelectorAll('a,button,div,span')]
        .find(e => e.innerText && e.innerText.trim() === '閉じる');
      if (btn) btn.click();
    });
    await new Promise(r => setTimeout(r, 1000));
    log('案内状況 更新完了');
    updateSheetStatus('R', `✅ ${nowJST()} 更新完了`);
  } catch (err) {
    log(`案内状況 エラー: ${err.message}`);
    updateSheetStatus('R', `❌ ${nowJST()} ${err.message}`);
    throw err;
  }
}

// ─────────────────────────────────────────
// リアルタイム集客
// ─────────────────────────────────────────
async function runRealtimeKyakka(page) {
  try {
    await page.goto('https://estama.jp/admin/', { waitUntil: 'networkidle2' });
    await ensureLoggedIn(page);
    await clickByText(page, 'リアルタイム集客');
    await safeWaitForNav(page);
    await clickByText(page, '新しいメッセージを書く');
    await safeWaitForNav(page);

    const templateResult = await page.evaluate((keyword) => {
      const options = [...document.querySelectorAll('option')].filter(e => e.value && e.value !== '');
      if (options.length === 0) return 'no_options';
      const target = keyword
        ? options.find(e => e.textContent.includes(keyword))
        : options.find(e => e.textContent.includes('新人入店記念キャンペーン'));
      const chosen = target || options[0];
      chosen.selected = true;
      chosen.closest('select').dispatchEvent(new Event('change', { bubbles: true }));
      return chosen.textContent.trim();
    }, store.templateKeyword);

    if (templateResult === 'no_options') {
      log('集客テンプレートなし スキップ');
      updateSheetStatus('S', `⏭️ ${nowJST()} テンプレートなしスキップ`);
      return;
    }
    log(`集客テンプレート: ${templateResult}`);
    await new Promise(r => setTimeout(r, 1500));

    const posted = await page.evaluate(() => {
      const btn = [...document.querySelectorAll('a,button,input[type="submit"]')]
        .find(e => e.innerText && e.innerText.includes('投稿'));
      if (btn) { btn.click(); return true; }
      return false;
    });
    if (!posted) {
      log('集客 投稿ボタンなし スキップ');
      updateSheetStatus('S', `⏭️ ${nowJST()} 投稿ボタンなしスキップ`);
      return;
    }
    await safeWaitForNav(page);
    log('集客 投稿完了');
    updateSheetStatus('S', `✅ ${nowJST()} 投稿完了`);
  } catch (err) {
    log(`集客 エラー: ${err.message}`);
    updateSheetStatus('S', `❌ ${nowJST()} ${err.message}`);
    throw err;
  }
}

// ─────────────────────────────────────────
// リアルタイム求人
// ─────────────────────────────────────────
async function runRealtimeKyujin(page) {
  try {
    await page.goto('https://estama.jp/admin/', { waitUntil: 'networkidle2' });
    await ensureLoggedIn(page);
    await clickByText(page, 'リアルタイム求人');
    await safeWaitForNav(page);

    const hasBtn = await page.evaluate(() => {
      const els = [...document.querySelectorAll('a,button,li,div,span')];
      return els.some(e => e.innerText && e.innerText.trim().includes('複製'));
    });
    if (!hasBtn) {
      log('求人 複製ボタンなし スキップ');
      updateSheetStatus('T', `⏭️ ${nowJST()} 複製ボタンなし（エスタマ管理画面で求人を1件手動投稿してください）`);
      return;
    }

    await clickByText(page, '複製');
    await safeWaitForNav(page);

    const posted = await page.evaluate(() => {
      const btn = [...document.querySelectorAll('a,button,input[type="submit"]')]
        .find(e => e.innerText && e.innerText.includes('投稿'));
      if (btn) { btn.click(); return true; }
      return false;
    });
    if (!posted) {
      log('求人 投稿ボタンなし スキップ');
      updateSheetStatus('T', `⏭️ ${nowJST()} 投稿ボタンなしスキップ`);
      return;
    }
    await safeWaitForNav(page);
    log('求人 投稿完了');
    updateSheetStatus('T', `✅ ${nowJST()} 投稿完了`);
  } catch (err) {
    log(`求人 エラー: ${err.message}`);
    updateSheetStatus('T', `❌ ${nowJST()} ${err.message}`);
    throw err;
  }
}

// ─────────────────────────────────────────
// 写メ日記 自動投稿
// ─────────────────────────────────────────
async function runShinshaPost(page, browser) {
  try {
    log('写メ日記投稿開始');

    // スケジュール取得 → 本文生成
    const staffList = await fetchTodaySchedule(browser, SHOP_ID);
    const body      = buildShinshaBody(staffList);

    // 新規投稿ページへ
    await page.goto('https://estama.jp/admin/blog/', { waitUntil: 'networkidle2' });
    await ensureLoggedIn(page);

    // 残り投稿数を確認
    const remaining = await page.evaluate(() => {
      const m = document.body.innerText.match(/本日の残り回数\s*(\d+)\s*回/);
      return m ? parseInt(m[1]) : -1;
    });
    if (remaining === 0) {
      log('写メ日記 本日の投稿上限に達しました スキップ');
      updateSheetStatus('U', `⏭️ ${nowJST()} 本日上限`);
      return;
    }

    // 「新しい写メ日記を書く」ボタンをクリック → ナビゲート待機
    await Promise.all([
      safeWaitForNav(page, 5000),
      page.evaluate(() => {
        const btn = [...document.querySelectorAll('a,button')]
          .find(e => e.innerText && e.innerText.includes('新しい写メ日記'));
        if (btn) btn.click();
      }),
    ]);

    // URL変わらない場合は直接アクセス
    if (!page.url().includes('blog_edit')) {
      await page.goto('https://estama.jp/admin/blog_edit/?r=temp', { waitUntil: 'networkidle2' });
    }
    log(`写メ日記フォームURL: ${page.url()}`);

    // タイトル入力
    await page.evaluate((title) => {
      const el = document.getElementById('PostTitle');
      if (el) { el.value = title; el.dispatchEvent(new Event('input', { bubbles: true })); }
    }, SHINSHA_TITLE);

    // 本文入力
    await page.evaluate((content) => {
      const el = document.getElementById('PostContent');
      if (el) { el.value = content; el.dispatchEvent(new Event('input', { bubbles: true })); }
    }, body);

    // 画像アップロード
    const imgPath = await downloadRandomImage(await getImageFolderId());
    await uploadImageToForm(page, imgPath);

    // 投稿ボタンをクリック
    const posted = await page.evaluate(() => {
      const btn = document.querySelector('a.send-post')
        || [...document.querySelectorAll('a,button')]
            .find(e => e.innerText && (e.innerText.includes('投稿') || e.innerText.includes('更新')));
      if (btn) { btn.click(); return true; }
      return false;
    });

    if (posted) {
      await safeWaitForNav(page, 10000);
      log('写メ日記 投稿完了');
      updateSheetStatus('U', `✅ ${nowJST()} 写メ日記投稿（${staffList.length}名）`);
    } else {
      log('写メ日記 投稿ボタンが見つかりません');
      updateSheetStatus('U', `❌ ${nowJST()} 投稿ボタンなし`);
    }
  } catch (err) {
    log(`写メ日記 エラー: ${err.message}`);
    updateSheetStatus('U', `❌ ${nowJST()} ${err.message}`);
    throw err;
  }
}

// ─────────────────────────────────────────
// 店長ブログ 自動投稿
// ─────────────────────────────────────────
async function runJobBlogPost(page) {
  try {
    log('店長ブログ投稿開始');

    // 最新投稿のコンテンツを取得
    await page.goto('https://estama.jp/admin/job_blog/', { waitUntil: 'networkidle2' });
    await ensureLoggedIn(page);

    // 残り投稿数を確認
    const remaining = await page.evaluate(() => {
      const m = document.body.innerText.match(/本日の残り回数\s*(\d+)\s*回/);
      return m ? parseInt(m[1]) : -1;
    });
    if (remaining === 0) {
      log('店長ブログ 本日の投稿上限に達しました スキップ');
      updateSheetStatus('V', `⏭️ ${nowJST()} 本日上限`);
      return;
    }

    // 最新の編集URLを取得してコンテンツをコピー
    // ※ 数字IDを持つ既存記事URLのみを対象とし「新しい求人ブログを書く」ボタンは除外する
    const latestEditUrl = await page.evaluate(() => {
      const links = [...document.querySelectorAll('a[href*="job_blog_edit"]')]
        .filter(a => /\/job_blog_edit\/\d+\//.test(a.href));
      return links.length > 0 ? links[0].href : null;
    });

    let blogContent = '';
    if (latestEditUrl) {
      await page.goto(latestEditUrl, { waitUntil: 'networkidle2' });
      blogContent = await page.evaluate(() => {
        const el = document.getElementById('PostContent');
        return el ? el.value : '';
      });
      log(`既存ブログ内容取得: ${blogContent.substring(0, 30)}...`);
      // ブログリストに戻る
      await page.goto('https://estama.jp/admin/job_blog/', { waitUntil: 'networkidle2' });
    }

    // 「新しい求人ブログを書く」ボタンをクリック
    await Promise.all([
      safeWaitForNav(page, 5000),
      page.evaluate(() => {
        const btn = [...document.querySelectorAll('a,button')]
          .find(e => e.innerText && e.innerText.includes('新しい求人ブログ'));
        if (btn) btn.click();
      }),
    ]);

    // URL変わらない場合は直接アクセス
    if (!page.url().includes('job_blog_edit')) {
      await page.goto('https://estama.jp/admin/job_blog_edit/', { waitUntil: 'networkidle2' });
    }
    log(`店長ブログフォームURL: ${page.url()}`);

    // タイトル入力
    await page.evaluate((title) => {
      const el = document.getElementById('PostTitle');
      if (el) { el.value = title; el.dispatchEvent(new Event('input', { bubbles: true })); }
    }, BLOG_TITLE);

    // 本文入力（BLOG_CONTENT優先、なければ既存記事コピー）
    const bodyToWrite = BLOG_CONTENT || blogContent;
    if (bodyToWrite) {
      await page.evaluate((content) => {
        const el = document.getElementById('PostContent');
        if (el) { el.value = content; el.dispatchEvent(new Event('input', { bubbles: true })); }
      }, bodyToWrite);
      log(`ブログ本文セット: ${bodyToWrite.substring(0, 20)}...`);
    }

    // 画像アップロード
    const imgPath = await downloadRandomImage(await getImageFolderId());
    await uploadImageToForm(page, imgPath);

    // 投稿ボタンをクリック
    const posted = await page.evaluate(() => {
      const btn = document.querySelector('a.send-post')
        || [...document.querySelectorAll('a,button')]
            .find(e => e.innerText && (e.innerText.includes('投稿') || e.innerText.includes('更新')));
      if (btn) { btn.click(); return true; }
      return false;
    });

    if (posted) {
      await safeWaitForNav(page, 10000);
      log('店長ブログ 投稿完了');
      updateSheetStatus('V', `✅ ${nowJST()} ブログ投稿`);
    } else {
      log('店長ブログ 投稿ボタンが見つかりません');
      updateSheetStatus('V', `❌ ${nowJST()} 投稿ボタンなし`);
    }
  } catch (err) {
    log(`店長ブログ エラー: ${err.message}`);
    updateSheetStatus('V', `❌ ${nowJST()} ${err.message}`);
    throw err;
  }
}

// ─────────────────────────────────────────
// ─────────────────────────────────────────
// 体験入店 掲載開始（1日1回）
// ─────────────────────────────────────────
async function runTaikenPost(page) {
  try {
    log('体験入店 掲載開始...');
    await page.goto('https://estama.jp/admin/', { waitUntil: 'networkidle2' });
    await ensureLoggedIn(page);
    await clickByText(page, '本日体験入店できるお店');
    await safeWaitForNav(page);

    const clicked = await page.evaluate(() => {
      const btn = [...document.querySelectorAll('a,button,input[type="submit"]')]
        .find(e => e.innerText && e.innerText.trim().includes('掲載開始'));
      if (btn) { btn.click(); return true; }
      return false;
    });

    if (!clicked) {
      log('体験入店 掲載開始ボタンなし（本日分は既に掲載済みの可能性）');
    } else {
      await safeWaitForNav(page);
      log('体験入店 掲載開始 完了');
    }
  } catch (err) {
    log(`体験入店 エラー: ${err.message}`);
  }
}

// ─────────────────────────────────────────
// 機能別ON/OFF設定（スプレッドシートから2分ごと更新）
// ─────────────────────────────────────────
let funcSettings = {
  annnai:  true,
  kyakka:  true,
  kyujin:  true,
  shinsha: SHINSHA_ENABLED,
  blog:    BLOG_ENABLED,
  taiken:  TAIKEN_ENABLED,
};
let lastFuncSettingsRefresh = 0;
const FUNC_SETTINGS_REFRESH_MS = 2 * 60 * 1000;

async function refreshFuncSettings() {
  if (Date.now() - lastFuncSettingsRefresh < FUNC_SETTINGS_REFRESH_MS) return;
  try {
    const sheets    = await getSheetsClient();
    const row       = await findStoreRow();
    if (!row) return;
    const sheetName = process.env.STORES_SHEET_NAME || '店舗設定';
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.STORES_SPREADSHEET_ID,
      range: `'${sheetName}'!W${row}:AB${row}`,
    });
    const vals = (res.data.values || [[]])[0] || [];
    const toB = (v, envDefault) => (v === undefined || v === '') ? envDefault : (v === 'TRUE' || v === true);
    funcSettings = {
      annnai:  toB(vals[0], true),
      kyakka:  toB(vals[1], true),
      kyujin:  toB(vals[2], true),
      shinsha: toB(vals[3], SHINSHA_ENABLED),
      blog:    toB(vals[4], BLOG_ENABLED),
      taiken:  toB(vals[5], TAIKEN_ENABLED),
    };
    lastFuncSettingsRefresh = Date.now();
  } catch (_) {}
}


// ─────────────────────────────────────────
// メイン
// ─────────────────────────────────────────
async function main() {
  log('起動');

  const launchArgs = [
    '--no-sandbox', '--disable-setuid-sandbox',
    '--disable-dev-shm-usage', '--disable-gpu',
    '--no-first-run', '--no-zygote', '--single-process',
  ];
  if (store.proxyHost && store.proxyPort) {
    launchArgs.push(`--proxy-server=${store.proxyHost}:${store.proxyPort}`);
  }

  const browser = await puppeteer.launch({
    headless: true,
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
    args: launchArgs,
  });

  const page = await browser.newPage();
  page.setDefaultNavigationTimeout(60000);
  await page.setViewport({ width: 1280, height: 900 });
  await page.setUserAgent('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36');

  if (store.proxyUser && store.proxyPass) {
    await page.authenticate({ username: store.proxyUser, password: store.proxyPass });
  }

  await login(page);

  let lastKyakkaRun = 0;
  let lastKyujinRun = 0;

  while (true) {
    const now = new Date();

    if (!isRunningTime(now)) {
      log(`停止中 (稼働時間外: ${STORE_START_TIME}〜${STORE_END_TIME})`);
      await new Promise(r => setTimeout(r, 60 * 1000));
      continue;
    }

    try {
      // スプレッドシートの機能ON/OFF設定を更新（2分ごと）
      await refreshFuncSettings();

      // 写メ日記・店長ブログ: 時刻ベースで投稿
      if (funcSettings.shinsha || funcSettings.blog) {
        const slotIndex = getDuePostSlotIndex();
        if (slotIndex >= 0) {
          postedSlotsToday.add(slotIndex);
          log(`投稿スロット #${slotIndex + 1}/10 実行`);
          if (funcSettings.shinsha) {
            try { await runShinshaPost(page, browser); } catch (e) { log(`写メ日記スロットエラー: ${e.message}`); }
          }
          if (funcSettings.blog) {
            try { await runJobBlogPost(page); } catch (e) { log(`ブログスロットエラー: ${e.message}`); }
          }
        }
      }

      // 体験入店：1日1回（日付が変わったら実行）
      if (funcSettings.taiken) {
        const { dateKey } = getJSTInfo();
        if (taikenDoneDate !== dateKey) {
          taikenDoneDate = dateKey;
          try { await runTaikenPost(page); } catch (e) { log(`体験入店エラー: ${e.message}`); }
        }
      }

      // 集客・求人（ON/OFFチェック）
      if (funcSettings.kyakka && Date.now() - lastKyakkaRun >= REALTIME_INTERVAL_MS) {
        await runRealtimeKyakka(page);
        lastKyakkaRun = Date.now();
      }
      if (funcSettings.kyujin && Date.now() - lastKyujinRun >= KYUJIN_INTERVAL_MS) {
        await runRealtimeKyujin(page);
        lastKyujinRun = Date.now();
      }
      if (funcSettings.annnai) {
        await runAnnaijokyo(page);
      }
    } catch (err) {
      log(`エラー: ${err.message}`);
      try {
        await ensureLoggedIn(page);
      } catch (e) {
        log(`セッション確認失敗: ${e.message}`);
      }
    }

    const jitter = Math.floor(Math.random() * 10 - 5) * 1000;
    await new Promise(r => setTimeout(r, ANNNAI_INTERVAL_MS + jitter));
  }
}

main().catch(err => {
  console.error(`[致命的エラー][${store.name}]`, err.message);
  process.exit(1);
});
