require('dotenv').config();
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const { getAvailability, getUpcomingSchedule, getDailyReportBookings, getIntervalMap, parseDateStr, dateToNippoName, warmupCache, getBusinessDate: sheetsGetBusinessDate } = require('./scheduleReader');
const { runEstamaSync } = require('./estama_worker');
const { lineConfig, handleLineEvent, middleware: lineMiddleware } = require('./line_handler');

const app = express();
app.use(express.static('public'));

// LINE webhook（署名検証なし・シンプル版でデバッグ）
app.post('/webhook/line', express.json(), (req, res) => {
  const events = req.body?.events || [];
  console.log('[LINE webhook] 受信 events:', events.length, JSON.stringify(events).slice(0, 300));
  events.forEach(ev => {
    if (ev.type === 'message' && ev.message?.type === 'text') {
      console.log('[LINE webhook] テキスト:', JSON.stringify(ev.message.text));
    }
  });
  res.status(200).end();
  Promise.all(events.map(handleLineEvent)).catch(err => console.error('[LINE webhook] エラー:', err.message));
});
console.log('LINE webhook 有効: POST /webhook/line');

app.use(express.json());

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// 空き状況データを読みやすいテキストにまとめる
function buildAvailabilityText(data) {
  const { results, now } = data;
  const lines = [`現在時刻: ${now}`, ''];

  const available = results.filter(r => r.status === '空き');
  const inService = results.filter(r => r.status === '接客中');
  const inInterval = results.filter(r => r.status === 'インターバル中');
  const closed = results.filter(r => r.status === '受付終了');
  const comingSoon = results.filter(r => r.status === '出勤前');
  // 退勤済み・本日出勤なしは表示しない

  if (available.length > 0) {
    lines.push('【今すぐ案内可能】');
    available.forEach(r => {
      if (r.hasUpper) {
        lines.push(`  - ${r.name}（最終受付: ${r.lastAccept}、${r.workEnd}終了）`);
      } else {
        lines.push(`  - ${r.name}（${r.workEnd}までにご来店で案内可能）`);
      }
    });
    lines.push('');
  }

  if (inService.length > 0) {
    lines.push('【接客中（新規案内不可）】');
    inService.forEach(r => lines.push(`  - ${r.name}（接客終了: ${r.serviceEnd}、案内可能: ${r.nextAvailable}〜）`));
    lines.push('');
  }

  if (inInterval.length > 0) {
    lines.push('【インターバル中】');
    inInterval.forEach(r => lines.push(`  - ${r.name}（案内可能: ${r.nextAvailable}〜）`));
    lines.push('');
  }

  if (closed.length > 0) {
    lines.push('【本日受付終了】');
    closed.forEach(r => lines.push(`  - ${r.name}（${r.note}）`));
    lines.push('');
  }

  if (comingSoon.length > 0) {
    lines.push('【まもなく出勤】');
    comingSoon.forEach(r => {
      const timeNote = r.hasUpper ? `${r.workEnd}終了` : `${r.workEnd}までにご来店で案内可能`;
      lines.push(`  - ${r.name}（出勤予定: ${r.nextAvailable}〜${timeNote}）`);
    });
  }

  return lines.join('\n');
}

const SYSTEM_PROMPT = `あなたはお店の出勤・予約案内専用アシスタントです。
入力ミスや言葉の省略があっても文脈から意図を読み取って答えてください。例えば「さくらもちは？」「沢田明日いる？」「27日さわだ空いてる？」のような質問も正しく理解して答えてください。


【絶対ルール】
- 出勤時間・空き状況・予約状況に関する質問にのみ答えてください
- それ以外（売上、給料、店の方針、雑談など）は「申し訳ございませんが、出勤・予約に関するご質問のみお答えしております」と答えてください

【予約ブロック時間の判断ルール】
- 【今日の予約状況と案内可否】に「○○〜○○は案内不可」と書かれている時間帯は絶対に案内できません
- お客様が「○時はどう？」と聞いてきた場合、その時間が案内不可の範囲内であれば「その時間帯は案内できません」と答えてください
- 「〜ご来店なら案内可」と「○○以降案内可」の間の時間帯は全て案内不可です

【今日の空き状況ルール】
- 「接客中」の子は絶対に「空き」と言わないでください。案内可能時間（nextAvailable）を伝えてください
- 「インターバル中」の子も案内可能時間を伝えてください
- 「受付終了」の子は本日受付終了と伝えてください
- 「空き」の子だけが今すぐ案内できます
- (上)あり → lastAcceptまでに来店必要、以降は案内不可
- (上)なし → workEndまでにご来店で案内可能
- 退勤済み・本日出勤なしの子は案内しないでください

【翌日以降のスケジュールルール】
- 全スタッフを省略せず全員の名前と時間帯をすべて表示（「など」「他〇名」禁止）
- (上)あり → 「〜終了」、(上)なし → 「〜までにご来店で案内可能」と表現
- データに「※満員」と書かれているスタッフは必ず「満員」と明記してください
- データに「※空きあり」と書かれているスタッフは空きがある旨を伝えてください
- 「※満員」「※空きあり」の記載がないスタッフは予約状況についての言及をしないでください

【予約状況ルール】
- 満員・案内不可の場合は「満員」とだけ伝えてください。予約の詳細（時間・件数など）は絶対に伝えないでください
- 予約がない場合は「現時点で予約はありません」と伝えてください

データにない情報は「わかりかねます」と伝えてください
回答にMarkdown記法（##、**、|---|などの表）は絶対に使わないでください。シンプルなテキストで答えてください。`;

const MIN_COURSE_MIN = 80;

// "HH:MM" を分に変換（25:00など24時超も対応）
function timeToMins(str) {
  if (!str) return 0;
  const m = String(str).match(/(\d+):(\d+)/);
  if (!m) return 0;
  return parseInt(m[1]) * 60 + parseInt(m[2]);
}

// 日またぎを考慮したbooking変換（end < start なら+1440）
function normalizeBooking(b, intervalMin) {
  const start = timeToMins(b.start);
  let end = timeToMins(b.end) + intervalMin;
  if (end < start) end += 1440;
  return { start, end };
}

// スタッフが満員かどうか判定（インターバルを考慮）
// currentTimeMins: 今日の場合は現在時刻（分）を渡す。過去の枠を除外するため
function isFullyBooked(staff, bookings, intervalMap = {}, currentTimeMins = null) {
  const intervalMin = intervalMap[staff.name] || 0;
  const myBookings = bookings
    .filter(b => b.name === staff.name)
    .map(b => normalizeBooking(b, intervalMin))
    .sort((a, b) => a.start - b.start);
  if (myBookings.length === 0) return false;

  const workStart = timeToMins(staff.start);
  const workEnd = timeToMins(staff.end);
  const lastAccept = staff.hasUpper ? workEnd - MIN_COURSE_MIN : workEnd;

  // 現在時刻が渡された場合、過去の枠はカウントしない
  let cursor = currentTimeMins !== null ? Math.max(workStart, currentTimeMins) : workStart;

  for (const b of myBookings) {
    if (b.start > cursor && b.start <= lastAccept) {
      if (b.start - cursor >= MIN_COURSE_MIN) return false; // 空きあり
    }
    if (b.end > cursor) cursor = b.end;
  }
  // 最終チェック: hasUpper=falseは残り1分以上あれば受付可（セッションが終了時刻を超えてもOK）
  const remaining = lastAccept - cursor;
  if (staff.hasUpper) {
    if (remaining >= MIN_COURSE_MIN) return false;
  } else {
    if (remaining > 0) return false;
  }
  return true;
}

// インターバル考慮後の次の案内可能時間を分で返す（予約がなければnull）
function getNextAvailableMins(staff, bookings, intervalMap = {}, currentTimeMins = null) {
  const intervalMin = intervalMap[staff.name] || 0;
  const myBookings = bookings
    .filter(b => b.name === staff.name)
    .map(b => normalizeBooking(b, intervalMin))
    .sort((a, b) => a.start - b.start);
  if (myBookings.length === 0) return null;

  const workStart = timeToMins(staff.start);
  let cursor = currentTimeMins !== null ? Math.max(workStart, currentTimeMins) : workStart;
  for (const b of myBookings) {
    if (b.end > cursor) cursor = b.end;
  }
  return cursor;
}

function minsToTimeStr(mins) {
  const h = Math.floor(mins / 60);
  const m = String(mins % 60).padStart(2, '0');
  return `${h}:${m}`;
}

// 営業日の基準日を返す（27時制：午前3時未満は前日扱い）
function getBusinessDate() {
  const now = new Date();
  if (now.getHours() < 3) {
    const d = new Date(now);
    d.setDate(d.getDate() - 1);
    return d;
  }
  return now;
}

// チャット履歴（メモリ内、シンプル実装）
const sessions = {};

app.post('/api/chat', async (req, res) => {
  const { message, sessionId } = req.body;
  if (!message) return res.status(400).json({ error: 'message is required' });

  const sid = sessionId || 'default';
  if (!sessions[sid]) sessions[sid] = [];

  try {
    // メッセージ中の日付を検出
    const datePatterns = [
      /明後日/, /明日/, /今日/, /本日/,
      /(\d{4})年(\d{1,2})月(\d{1,2})日/,
      /(\d{1,2})月(\d{1,2})日/,
      /(\d{1,2})日/,
    ];
    let mentionedDate = null;
    for (const pat of datePatterns) {
      const m = message.match(pat);
      if (m) { mentionedDate = m[0]; break; }
    }

    // 今日のリアルタイムデータ＋翌日以降スケジュール＋インターバルマップ＋今日の日報を並行取得
    const businessDate = getBusinessDate();
    const todayNippoName = dateToNippoName(businessDate);
    const [availability, upcoming, intervalMapResult, todayNippo] = await Promise.all([
      getAvailability(),
      getUpcomingSchedule(),
      getIntervalMap().catch(e => { console.error('インターバルマップ取得エラー:', e.message); return {}; }),
      getDailyReportBookings(todayNippoName).catch(e => { console.error('今日の日報取得エラー:', e.message); return { bookings: [] }; }),
    ]);
    const intervalMap = intervalMapResult;
    const availText = buildAvailabilityText(availability);

    // 今日の予約状況テキストを構築（常にコンテキストに含める）
    const todayStaff = availability.results.filter(r => r.start && r.end);
    let todayBookingText = `\n\n【今日（${todayNippoName}）の予約状況と案内可否】\n`;
    if (todayNippo.bookings.length === 0) {
      todayBookingText += '予約データなし（全員案内可能）\n';
    } else {
      for (const s of todayStaff) {
        const myBookings = todayNippo.bookings.filter(b => b.name === s.name);
        if (myBookings.length === 0) {
          todayBookingText += `  - ${s.name}: 【予約なし・案内可能】\n`;
          continue;
        }
        const rawNowMins = timeToMins(availability.now);
        // 深夜0〜2時は営業日の延長（24:xx扱い）
        const nowMins = (new Date()).getHours() < 3 ? rawNowMins + 1440 : rawNowMins;
        const full = isFullyBooked(s, todayNippo.bookings, intervalMap, nowMins);
        const nextMins = getNextAvailableMins(s, todayNippo.bookings, intervalMap, nowMins);
        const intervalMin = intervalMap[s.name] || 0;

        if (full) {
          todayBookingText += `  - ${s.name}: 【満員・本日案内不可】\n`;
        } else {
          // 最初の予約開始時刻を取得
          const sortedBookings = myBookings
            .map(b => normalizeBooking(b, intervalMin))
            .sort((a, b) => a.start - b.start);
          const firstBookingStart = sortedBookings[0].start;
          const lastSafeArrival = firstBookingStart - MIN_COURSE_MIN;

          // 上あり: lastAccept = workEnd - 80分、上なし: workEnd まで
          const workEndMins = timeToMins(s.end);
          const lastAcceptMins = s.hasUpper ? workEndMins - MIN_COURSE_MIN : workEndMins;

          // nextMinsがlastAcceptを超えていれば予約後の枠はない
          const hasAfterSlot = nextMins !== null && nextMins <= lastAcceptMins;

          const parts = [];
          if (lastSafeArrival > timeToMins(s.start)) {
            parts.push(`〜${minsToTimeStr(lastSafeArrival)}ご来店なら案内可`);
          }
          if (hasAfterSlot) {
            parts.push(`${minsToTimeStr(nextMins)}以降案内可`);
            parts.push(`${minsToTimeStr(lastSafeArrival + 1)}〜${minsToTimeStr(nextMins - 1)}は案内不可`);
          } else {
            parts.push(`${minsToTimeStr(lastSafeArrival + 1)}以降は案内不可`);
          }
          todayBookingText += `  - ${s.name}: 【予約あり】${parts.join('、')}\n`;
        }
      }
    }

    // 日付が言及された場合は日報＋満員判定を先に実施
    const fullMapForDate = {};        // scheduleKey → { name → 満員boolean }
    const nextAvailMapForDate = {};   // scheduleKey → { name → 次案内可能時間（分） }
    const lastSafeMapForDate = {};    // scheduleKey → { name → 予約前の最終受付時刻（分） }
    if (mentionedDate) {
      try {
        const targetDate = parseDateStr(mentionedDate, businessDate);
        if (targetDate) {
          const nippoName = dateToNippoName(targetDate);
          const scheduleKey = `${targetDate.getMonth() + 1}月${targetDate.getDate()}日`;
          const isToday = targetDate.toDateString() === businessDate.toDateString();
          const nippo = isToday ? todayNippo : await getDailyReportBookings(nippoName);
          const daySchedule = isToday
            ? availability.results.filter(r => r.start && r.end)
            : (upcoming[scheduleKey] || []);
          const fm = {};
          const nav = {};
          const lsa = {};
          for (const s of daySchedule) {
            fm[s.name] = isFullyBooked(s, nippo.bookings, intervalMap);
            const nextMins = getNextAvailableMins(s, nippo.bookings, intervalMap);
            const intervalMin = intervalMap[s.name] || 0;
            const myBookings = nippo.bookings
              .filter(b => b.name === s.name)
              .map(b => normalizeBooking(b, intervalMin))
              .sort((a, b) => a.start - b.start);
            if (myBookings.length > 0) {
              lsa[s.name] = myBookings[0].start - MIN_COURSE_MIN; // 予約前の最終受付（分）
            }
            // lastAcceptを超えるnextMinsは無効
            const workEndMins = timeToMins(s.end);
            const lastAcceptMins = s.hasUpper ? workEndMins - MIN_COURSE_MIN : workEndMins;
            if (nextMins !== null && nextMins <= lastAcceptMins) {
              nav[s.name] = nextMins;
            }
          }
          fullMapForDate[scheduleKey] = fm;
          nextAvailMapForDate[scheduleKey] = nav;
          lastSafeMapForDate[scheduleKey] = lsa;
        }
      } catch (e) {
        console.error('日報取得エラー:', e.message);
      }
    }

    // 翌日以降スケジュールをテキスト化（満員情報・次案内可能時間を統合）
    let upcomingText = '';
    const upcomingDays = Object.entries(upcoming);
    if (upcomingDays.length > 0) {
      upcomingText = '\n\n【翌日以降の出勤スケジュール】\n';
      for (const [date, staff] of upcomingDays) {
        upcomingText += `\n${date}:\n`;
        const fm = fullMapForDate[date] || {};
        const nav = nextAvailMapForDate[date] || {};
        const lsa = lastSafeMapForDate[date] || {};
        staff.forEach(s => {
          const timeNote = s.hasUpper ? `${s.end}終了` : `${s.end}までにご来店で案内可能`;
          let status = '';
          if (fm[s.name] === true) {
            status = ' ※満員';
          } else if (lsa[s.name] !== undefined || nav[s.name] !== undefined) {
            const parts = [];
            if (lsa[s.name] !== undefined) {
              parts.push(`〜${minsToTimeStr(lsa[s.name])}ご来店なら案内可`);
            }
            if (nav[s.name] !== undefined) {
              parts.push(`${minsToTimeStr(nav[s.name])}以降案内可`);
            }
            if (lsa[s.name] !== undefined && nav[s.name] !== undefined) {
              parts.push(`${minsToTimeStr(lsa[s.name] + 1)}〜${minsToTimeStr(nav[s.name] - 1)}は案内不可`);
            } else if (lsa[s.name] !== undefined && nav[s.name] === undefined) {
              parts.push(`${minsToTimeStr(lsa[s.name] + 1)}以降は案内不可`);
            }
            status = ' ※予約あり（' + parts.join('、') + '）';
          }
          upcomingText += `  - ${s.name}: ${s.start}〜${timeNote}${status}\n`;
        });
      }
    }

    const userContent = `【現在の空き時間データ（${availability.now}時点）】\n${availText}${todayBookingText}${upcomingText}\n\n【お客様の質問】\n${message}`;

    sessions[sid].push({ role: 'user', content: userContent });

    const response = await anthropic.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 2048,
      system: SYSTEM_PROMPT,
      messages: sessions[sid],
    });

    const reply = response.content[0].text;
    sessions[sid].push({ role: 'assistant', content: reply });

    // セッションが長くなりすぎないよう最新20件に制限
    if (sessions[sid].length > 20) {
      sessions[sid] = sessions[sid].slice(-20);
    }

    res.json({ reply });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// 空き状況を直接取得するエンドポイント（デバッグ用）
app.get('/api/availability', async (req, res) => {
  try {
    const data = await getAvailability();
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// デバッグ: 指定日の日報と満員判定を確認
// 例: GET /api/debug/3月27日
app.get('/api/debug/:date', async (req, res) => {
  try {
    const dateStr = decodeURIComponent(req.params.date);
    const targetDate = parseDateStr(dateStr, getBusinessDate());
    if (!targetDate) return res.json({ error: '日付解析失敗', input: dateStr });

    const nippoName = dateToNippoName(targetDate);
    const scheduleKey = `${targetDate.getMonth() + 1}月${targetDate.getDate()}日`;
    const [upcoming, nippo, intervalMap] = await Promise.all([getUpcomingSchedule(), getDailyReportBookings(nippoName), getIntervalMap()]);

    const daySchedule = upcoming[scheduleKey] || [];
    const fm = {};
    const nav = {};
    for (const s of daySchedule) {
      fm[s.name] = isFullyBooked(s, nippo.bookings, intervalMap);
      const nextMins = getNextAvailableMins(s, nippo.bookings, intervalMap);
      if (nextMins !== null) nav[s.name] = minsToTimeStr(nextMins);
    }

    res.json({
      scheduleKey,
      nippoName,
      daySchedule,
      nippoBookings: nippo.bookings,
      intervalMap,
      fullyBookedMap: fm,
      nextAvailableMap: nav,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 新しいチャットセッション開始
app.post('/api/reset', (req, res) => {
  const { sessionId } = req.body;
  const sid = sessionId || 'default';
  sessions[sid] = [];
  res.json({ ok: true });
});

// estama出勤表 × 更新（SSEでリアルタイムログ配信）
app.get('/api/estama-sync', async (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const send = (msg) => {
    res.write(`data: ${JSON.stringify({ message: msg })}\n\n`);
  };

  try {
    const result = await runEstamaSync(send);
    res.write(`data: ${JSON.stringify({ done: true, ...result })}\n\n`);
  } catch (e) {
    res.write(`data: ${JSON.stringify({ error: e.message })}\n\n`);
  }
  res.end();
});

// デバッグ: sharp動作確認 + 手動でPush送信をトリガー
app.get('/api/test-line', async (req, res) => {
  const steps = [];
  try {
    steps.push('1. sharp テスト');
    const sharp = require('sharp');
    const fs = require('fs'), path = require('path');
    const tmpDir = path.join(__dirname, 'public', 'temp');
    fs.mkdirSync(tmpDir, { recursive: true });
    const tmpFile = path.join(tmpDir, `test_${Date.now()}.png`);
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="30"><rect width="100" height="30" fill="#fff"/><text x="5" y="20" font-size="12">テスト</text></svg>';
    await sharp(Buffer.from(svg)).png().toFile(tmpFile);
    const stat = fs.statSync(tmpFile);
    steps.push(`sharp OK: ${stat.size} bytes`);

    steps.push('2. LINE Push テスト');
    const { messagingApi } = require('@line/bot-sdk');
    const client = new messagingApi.MessagingApiClient({ channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN });
    const imageUrl = `${process.env.LINE_BOT_SERVER_URL}/temp/${path.basename(tmpFile)}`;
    await client.pushMessage({
      to: process.env.LINE_USER_ID,
      messages: [{ type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl }],
    });
    steps.push(`Push OK: ${imageUrl}`);

    res.json({ ok: true, steps });
  } catch (e) {
    steps.push(`ERROR: ${e.message}`);
    res.status(500).json({ ok: false, steps, error: e.message, stack: e.stack?.split('\n').slice(0,5) });
  }
});

// デバッグ: 手動で (教えて) フロー全体を実行
app.get('/api/test-line-full', async (req, res) => {
  res.json({ ok: true, message: 'バックグラウンドでスクショ処理開始' });
  const { messagingApi } = require('@line/bot-sdk');
  const client = new messagingApi.MessagingApiClient({ channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN });
  const target = process.env.LINE_USER_ID;
  setImmediate(async () => {
    try {
      const { findSpreadsheetId, screenshotCells, getSheetBusinessDate } = require('./line_handler_test');
    } catch(e) {}
    // line_handler の processAndPush と同じ処理をここで直接実行
    try {
      const { google } = require('googleapis');
      const sharp = require('sharp');
      const fs = require('fs'), path = require('path');

      const now = new Date();
      const d = now.getHours() < 6 ? new Date(now.getTime() - 86400000) : now;
      const dateStr = `${d.getFullYear()}年${d.getMonth()+1}月${d.getDate()}日`;
      console.log('[TEST] 検索日:', dateStr);

      const auth = (() => {
        const client_id = process.env.GOOGLE_CLIENT_ID;
        const client_secret = process.env.GOOGLE_CLIENT_SECRET;
        const refresh_token = process.env.GOOGLE_REFRESH_TOKEN;
        const c = new google.auth.OAuth2(client_id, client_secret, 'http://localhost');
        c.setCredentials({ refresh_token });
        return c;
      })();

      const drive = google.drive({ version: 'v3', auth });
      const res2 = await drive.files.list({
        q: `'1isPYyiUqyWXnS1mtpE1_YWJ9QZBTemdJ' in parents and name contains '${dateStr}'`,
        fields: 'files(id,name)', pageSize: 5,
      });
      const file = res2.data.files?.[0];
      if (!file) throw new Error(`ファイルなし: ${dateStr}`);
      console.log('[TEST] ファイル:', file.name);

      const sheets = google.sheets({ version: 'v4', auth });
      const meta = await sheets.spreadsheets.get({ spreadsheetId: file.id });
      const sheet = meta.data.sheets.find(s => s.properties.sheetId === 1873674341);
      console.log('[TEST] シート:', sheet?.properties?.title);

      const sheetData = await sheets.spreadsheets.get({
        spreadsheetId: file.id,
        ranges: [`'${sheet.properties.title}'!CA3:CR32`],
        includeGridData: true,
      });
      const gridData = sheetData.data.sheets?.[0]?.data?.[0];
      const rows = gridData?.rowData || [];
      console.log('[TEST] 行数:', rows.length);

      const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="50"><rect width="200" height="50" fill="#fff" stroke="#ccc"/><text x="5" y="30" font-size="14">取得OK: ' + rows.length + '行</text></svg>';
      const tmpDir = path.join(__dirname, 'public', 'temp');
      fs.mkdirSync(tmpDir, { recursive: true });
      const tmpFile = path.join(tmpDir, `full_${Date.now()}.png`);
      await sharp(Buffer.from(svg)).png().toFile(tmpFile);
      console.log('[TEST] sharp OK');

      const { messagingApi: ma } = require('@line/bot-sdk');
      const c2 = new ma.MessagingApiClient({ channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN });
      const imageUrl = `${process.env.LINE_BOT_SERVER_URL}/temp/${path.basename(tmpFile)}`;
      await c2.pushMessage({ to: target, messages: [{ type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl }] });
      console.log('[TEST] Push OK:', imageUrl);
    } catch(e) {
      console.error('[TEST] エラー:', e.message);
      try {
        const { messagingApi: ma } = require('@line/bot-sdk');
        const c2 = new ma.MessagingApiClient({ channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN });
        await c2.pushMessage({ to: target, messages: [{ type: 'text', text: `[デバッグ]エラー: ${e.message}` }] });
      } catch(_) {}
    }
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`サーバー起動: http://localhost:${PORT}`);
  warmupCache().catch(err => console.error('ウォームアップエラー:', err.message));
});
