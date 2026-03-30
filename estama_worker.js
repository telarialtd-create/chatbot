/**
 * estama_worker.js
 * esatama出勤表 × 更新のコアロジック（モジュール）
 */
require('dotenv').config();
const puppeteer = require('puppeteer');
const { getDailyReportBookings, getIntervalMap, getBusinessDate, dateToNippoName } = require('./scheduleReader');

const ESTAMA_SCHEDULE_URL = 'https://estama.jp/admin/schedule/';

function timeToMins(str) {
  if (!str) return null;
  const m = String(str).match(/(\d+):(\d+)/);
  if (!m) return null;
  return parseInt(m[1]) * 60 + parseInt(m[2]);
}

function minsToTime(mins) {
  const h = Math.floor(mins / 60);
  const m = String(mins % 60).padStart(2, '0');
  return `${h}:${m}`;
}

function getTodayStr() {
  const biz = getBusinessDate();
  const y = biz.getFullYear();
  const mo = String(biz.getMonth() + 1).padStart(2, '0');
  const d = String(biz.getDate()).padStart(2, '0');
  return `${y}-${mo}-${d}`;
}

function calcBlockedSlots(start, end, intervalMin = 0) {
  const startMins = timeToMins(start);
  if (startMins === null) return [];
  let endMins = (end && timeToMins(end) !== null) ? timeToMins(end) : startMins + 80;
  endMins += intervalMin;
  const slotStart = Math.floor(startMins / 30) * 30;
  const slotEnd = Math.ceil(endMins / 30) * 30;
  const slots = [];
  for (let t = slotStart; t < slotEnd; t += 30) slots.push(minsToTime(t));
  return slots;
}

function findCastId(shortName, staffIdMap) {
  if (staffIdMap[shortName]) return staffIdMap[shortName];
  for (const [fullName, castId] of Object.entries(staffIdMap)) {
    if (fullName.includes(shortName) || shortName.includes(fullName)) return castId;
  }
  return null;
}

/**
 * メイン実行関数
 * @param {function} log - ログ出力コールバック (message: string) => void
 */
async function runEstamaSync(log = console.log) {
  const todayStr = getTodayStr();
  const bizDate = getBusinessDate();
  const nippoName = dateToNippoName(bizDate);

  log(`対象日: ${todayStr} (${nippoName})`);

  // スプレッドシートから予約取得
  log('スプレッドシートから予約を取得中...');
  let nippoData, intervalMap;
  try {
    [nippoData, intervalMap] = await Promise.all([
      getDailyReportBookings(nippoName),
      getIntervalMap(),
    ]);
  } catch (e) {
    log(`エラー: ${e.message}`);
    throw e;
  }

  if (!nippoData.bookings || nippoData.bookings.length === 0) {
    log('本日の予約なし。処理を終了します。');
    return { success: true, processed: 0, skipped: 0 };
  }

  log(`予約件数: ${nippoData.bookings.length}件`);

  // スタッフ別ブロックスロット計算
  const staffSlots = {};
  for (const booking of nippoData.bookings) {
    const interval = intervalMap[booking.name] || 0;
    const blocked = calcBlockedSlots(booking.start, booking.end, interval);
    if (!blocked.length) continue;
    if (!staffSlots[booking.name]) staffSlots[booking.name] = new Set();
    blocked.forEach(s => staffSlots[booking.name].add(s));
  }

  // Puppeteerでestama操作
  log('estama.jpにログイン中...');
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 900 });

  await page.goto('https://estama.jp/login/?r=/admin/', { waitUntil: 'networkidle2' });
  await page.type('#inputEmail', process.env.ESTAMA_USER, { delay: 30 });
  await page.type('#inputPassword', process.env.ESTAMA_PASS, { delay: 30 });
  try {
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }),
      page.click('a[data-post="login_shop"]'),
    ]);
  } catch (e) {
    await browser.close();
    throw new Error('ログイン失敗: ' + e.message);
  }
  log('ログイン成功');

  // スタッフID取得
  await page.goto(ESTAMA_SCHEDULE_URL, { waitUntil: 'networkidle2' });
  const staffIdMap = await page.evaluate(() => {
    const map = {};
    document.querySelectorAll('th.l-work_schedule_cell[data-idx]').forEach(th => {
      const castId = th.getAttribute('data-idx');
      const nameEl = th.querySelector('span');
      if (castId && nameEl) map[nameEl.innerText.trim()] = castId;
    });
    return map;
  });

  let processed = 0, skipped = 0, totalChanged = 0;

  for (const [name, slotsSet] of Object.entries(staffSlots)) {
    const castId = findCastId(name, staffIdMap);
    if (!castId) {
      log(`スキップ: ${name}（estama未登録）`);
      skipped++;
      continue;
    }

    const slots = [...slotsSet].sort();
    log(`処理中: ${name} → ${slots.join(', ')}`);

    await page.goto(`${ESTAMA_SCHEDULE_URL}${castId}/`, { waitUntil: 'networkidle2' });

    let changed = 0;
    for (const slot of slots) {
      const selector = `select[name="column[${todayStr}][period][${slot}]"]`;
      try {
        const el = await page.$(selector);
        if (!el) continue;
        const val = await page.$eval(selector, e => e.value);
        if (val === '2') continue;
        await page.select(selector, '2');
        changed++;
      } catch {}
    }

    if (changed > 0) {
      await page.click('#SendWorkSchedule');
      await new Promise(r => setTimeout(r, 2000));
      log(`  → ${changed}スロット × に設定・保存完了`);
      totalChanged += changed;
    } else {
      log(`  → 変更なし（設定済み）`);
    }
    processed++;
  }

  await browser.close();
  log(`完了: ${processed}名処理, ${totalChanged}スロット更新, ${skipped}名スキップ`);
  return { success: true, processed, skipped, totalChanged };
}

module.exports = { runEstamaSync };
