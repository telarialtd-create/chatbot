require('dotenv').config();
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const { getAvailability, getUpcomingSchedule, getDailyReportBookings, getIntervalMap, parseDateStr, dateToNippoName } = require('./sheets');

const app = express();
app.use(express.json());
app.use(express.static('public'));

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
    comingSoon.forEach(r => lines.push(`  - ${r.name}（出勤予定: ${r.nextAvailable}〜）`));
  }

  return lines.join('\n');
}

const SYSTEM_PROMPT = `あなたはお店の出勤・予約案内専用アシスタントです。
入力ミスや言葉の省略があっても文脈から意図を読み取って答えてください。例えば「さくらもちは？」「沢田明日いる？」「27日さわだ空いてる？」のような質問も正しく理解して答えてください。


【絶対ルール】
- 出勤時間・空き状況・予約状況に関する質問にのみ答えてください
- それ以外（売上、給料、店の方針、雑談など）は「申し訳ございませんが、出勤・予約に関するご質問のみお答えしております」と答えてください

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
- 「※満員」「※空きあり」の記載がないスタッフは予約状況不明として扱ってください

【予約状況ルール】
- 日報データに予約がある場合は名前・開始時間・終了時間を伝えてください
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

// スタッフが満員かどうか判定（インターバルを考慮）
function isFullyBooked(staff, bookings, intervalMap = {}) {
  const intervalMin = intervalMap[staff.name] || 0;
  const myBookings = bookings
    .filter(b => b.name === staff.name)
    .map(b => ({ start: timeToMins(b.start), end: timeToMins(b.end) + intervalMin }))
    .sort((a, b) => a.start - b.start);
  if (myBookings.length === 0) return false;

  const workStart = timeToMins(staff.start);
  const workEnd = timeToMins(staff.end);
  // (上)あり: lastAccept = workEnd - 80分、(上)なし: workEnd まで受付可
  const lastAccept = staff.hasUpper ? workEnd - MIN_COURSE_MIN : workEnd;

  // workStart〜lastAccept の間に80分以上の空きがあるか確認
  let cursor = workStart;
  for (const b of myBookings) {
    if (b.start > cursor && b.start <= lastAccept) {
      if (b.start - cursor >= MIN_COURSE_MIN) return false; // 空きあり
    }
    if (b.end > cursor) cursor = b.end;
  }
  // 最後の予約以降
  if (lastAccept - cursor >= MIN_COURSE_MIN) return false;
  return true;
}

// インターバル考慮後の次の案内可能時間を分で返す（予約がなければnull）
function getNextAvailableMins(staff, bookings, intervalMap = {}) {
  const intervalMin = intervalMap[staff.name] || 0;
  const myBookings = bookings
    .filter(b => b.name === staff.name)
    .map(b => ({ start: timeToMins(b.start), end: timeToMins(b.end) + intervalMin }))
    .sort((a, b) => a.start - b.start);
  if (myBookings.length === 0) return null;

  const workStart = timeToMins(staff.start);
  let cursor = workStart;
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
      /明後日/, /明日/, /今日/,
      /(\d{4})年(\d{1,2})月(\d{1,2})日/,
      /(\d{1,2})月(\d{1,2})日/,
    ];
    let mentionedDate = null;
    for (const pat of datePatterns) {
      const m = message.match(pat);
      if (m) { mentionedDate = m[0]; break; }
    }

    // 今日のリアルタイムデータ＋翌日以降スケジュールを取得
    const [availability, upcoming] = await Promise.all([getAvailability(), getUpcomingSchedule()]);
    const availText = buildAvailabilityText(availability);

    // インターバルマップを取得
    let intervalMap = {};
    try {
      intervalMap = await getIntervalMap();
    } catch (e) {
      console.error('インターバルマップ取得エラー:', e.message);
    }

    // 日付が言及された場合は日報＋満員判定を先に実施
    const fullMapForDate = {}; // scheduleKey → { name → 満員boolean }
    const nextAvailMapForDate = {}; // scheduleKey → { name → 次案内可能時間文字列 }
    if (mentionedDate) {
      try {
        const targetDate = parseDateStr(mentionedDate);
        const today = new Date();
        if (targetDate && targetDate.toDateString() !== today.toDateString()) {
          const nippoName = dateToNippoName(targetDate);
          const scheduleKey = `${targetDate.getMonth() + 1}月${targetDate.getDate()}日`;
          const nippo = await getDailyReportBookings(nippoName);
          const daySchedule = upcoming[scheduleKey] || [];
          const fm = {};
          const nav = {};
          for (const s of daySchedule) {
            fm[s.name] = isFullyBooked(s, nippo.bookings, intervalMap);
            const nextMins = getNextAvailableMins(s, nippo.bookings, intervalMap);
            if (nextMins !== null) nav[s.name] = minsToTimeStr(nextMins);
          }
          fullMapForDate[scheduleKey] = fm;
          nextAvailMapForDate[scheduleKey] = nav;
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
        staff.forEach(s => {
          const timeNote = s.hasUpper ? `${s.end}終了` : `${s.end}までにご来店で案内可能`;
          let status = '';
          if (fm[s.name] === true) {
            status = ' ※満員';
          } else if (nav[s.name]) {
            status = ` ※${nav[s.name]}〜案内可能（現在予約あり）`;
          }
          upcomingText += `  - ${s.name}: ${s.start}〜${timeNote}${status}\n`;
        });
      }
    }

    const nippoText = '';

    const userContent = `【現在の空き時間データ（${availability.now}時点）】\n${availText}${upcomingText}${nippoText}\n\n【お客様の質問】\n${message}`;

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
    const targetDate = parseDateStr(dateStr);
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`サーバー起動: http://localhost:${PORT}`);
});
