/**
 * shift_reflect.js
 *
 * LINE出勤表グループに投稿されたシフトメッセージを、月次シフトSSに直接反映。
 * C-002方式（GAS Web App非経由、サービスアカウントで直接書き込み）。
 *
 * 入力メッセージ形式（例）:
 *   セラピスト名
 *   22日(水)休み
 *   23日(木)19~3受け
 *   24日(金)20~3受け
 *   1日(金)19~3受け   ← 日が戻ったら翌月扱い
 *
 * SS前提:
 *   ・タイトル: "YYYY年M月 シフト表"
 *   ・タブ: "CREA" / "ふわもこ" / "👥 スタッフ管理"
 *   ・CREA/ふわもこ行構成:
 *       row1: タイトル
 *       row2: ヘッダ（A="スタッフ名", B=当月1日 ... AE=当月30/31日,
 *              AF以降=翌月プレビュー1〜7日, 末尾に集計列）
 *       row3以降: スタッフ1人1行
 */

const { google } = require('googleapis');
const path = require('path');

let _auth = null;
function createAuth() {
  if (_auth) return _auth;
  _auth = new google.auth.GoogleAuth({
    keyFile: path.join(process.env.HOME, '.config/chatbot-service-account.json'),
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
    ],
  });
  return _auth;
}

// 全角数字 → 半角
function toHalfDigits(s) {
  return String(s || '').replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
}

// 空白のみ除去した正規化（数字/英字は残す。"144　峯岸みなみ" のように数字メモが混じる名前も許容）
function normalizeName(s) {
  return String(s || '').replace(/[\s　]+/g, '');
}

// 列番号（1-based）→ 列文字
function colIndexToLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function daysInMonth(year, month /* 1-based */) {
  return new Date(year, month, 0).getDate();
}

function weekdayKanji(year, month, day) {
  const d = new Date(Date.UTC(year, month - 1, day));
  return '日月火水木金土'[d.getUTCDay()];
}

// JST now
function getJSTDate() {
  return new Date(Date.now() + 9 * 60 * 60 * 1000);
}

// ── パース ─────────────────────────────────────────────

/**
 * シフトメッセージを解析（v2: 抽出方式 / 1行投稿 / 範囲指定 / 文章中日付 対応）
 */
function parseShiftMessage(text) {
  // v5.2追加: 「今日」を JST投稿日の「N日(曜)今日」に展開（"今日"自体は残す→normalize側で当欠検出）
  // 1行目(スタッフ名)は除外して、2行目以降のみ展開する
  let _expanded = String(text || '');
  {
    const _t = getJSTDate();
    const _replacement = `${_t.getDate()}日(${'日月火水木金土'[_t.getDay()]})今日`;
    const _parts = _expanded.split(/\r?\n/);
    for (let i = 1; i < _parts.length; i++) {
      _parts[i] = _parts[i].replace(/今日/g, _replacement);
    }
    _expanded = _parts.join('\n');
  }

  let lines = _expanded
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(l => l.length > 0);

  if (lines.length === 0) return null;

  // 完全1行投稿: 「みみ 4日(月)18-24受」「みみ 4(月)18-24受」
  if (lines.length === 1) {
    const line = toHalfDigits(lines[0]);
    const m = line.match(/^(.+?)[\s　]+(\d{1,2}(?:日|(?=[\(（]\s*[月火水木金土日])).*)$/);
    if (m && !/^\d+(?:日|[\(（])/.test(toHalfDigits(m[1].trim()))) {
      lines = [m[1].trim(), m[2].trim()];
    } else {
      return null;
    }
  }

  if (lines.length < 2) return null;

  const name = lines[0];
  if (/^\d+(?:日|[\(（])/.test(toHalfDigits(name))) return null;

  const entries = [];
  // 範囲: 「1日(金)〜6日(水・振替休日) おやすみ」「1〜6 おやすみ」「11(月)〜13(水) お休み」
  // v5.6: 「日」をoptional化（B案）
  const rangeRe = /(\d{1,2})(?:日)?(?:[\(（][^\)）]*[\)）])?\s*[〜～~ー]\s*(\d{1,2})(?:日)?(?:[\(（][^\)）]*[\)）])?\s*(.*)/;
  // 単発日付: 「\d+日(曜日?)」「\d+(曜日)」 ※(月) (月・祝) (水・振替休日) を許容
  // v5.6: 「日」 or 「曜日カッコ前」のいずれかを必須化（B案）
  // 直前が数字/コロン/ドットの場合は除外（時刻の部分を日付と誤認しないため）
  const dayTokenRe = /(?<![:.\d])(\d{1,2})(?:日|(?=[\(（]\s*[月火水木金土日]))\s*(?:[\(（]\s*([月火水木金土日])(?:[・,、][^\)）]*)?\s*[\)）])?/g;

  let pending = null;
  for (let i = 1; i < lines.length; i++) {
    const raw = lines[i];
    const line = toHalfDigits(raw);

    // 範囲指定を先に試す
    const rm = line.match(rangeRe);
    if (rm) {
      const sd = parseInt(rm[1], 10);
      const ed = parseInt(rm[2], 10);
      const value = (rm[3] || '').trim();
      if (sd >= 1 && sd <= 31 && ed >= 1 && ed <= 31 && value && sd <= ed) {
        for (let d = sd; d <= ed; d++) {
          entries.push({ day: d, weekday: null, value, raw });
        }
        pending = null;
        continue;
      }
    }

    // 行内のすべての「\d+日」トークンを集める
    const tokens = [];
    let dm;
    dayTokenRe.lastIndex = 0;
    while ((dm = dayTokenRe.exec(line)) !== null) {
      tokens.push({
        day: parseInt(dm[1], 10),
        weekday: dm[2] || null,
        start: dm.index,
        end: dayTokenRe.lastIndex,
      });
    }

    // v5.6: B案フォールバック - 行頭の単独「N」を日付候補として救済
    // 「15 18-2上」「15休み」のような「日」省略パターンを拾う
    // 直後が `:` `.` 数字 の場合は除外（時刻と誤認しないため）
    if (tokens.length === 0) {
      const headM = line.match(/^(\d{1,2})(?![:.\d])/);
      if (headM) {
        const day = parseInt(headM[1], 10);
        if (day >= 1 && day <= 31) {
          tokens.push({
            day,
            weekday: null,
            start: 0,
            end: headM[0].length,
          });
        }
      }
    }

    if (tokens.length === 0) {
      if (pending) {
        entries.push({ day: pending.day, weekday: pending.weekday, value: raw, raw: pending.raw + ' / ' + raw });
        pending = null;
      }
      continue;
    }

    pending = null;

    for (let j = 0; j < tokens.length; j++) {
      const tok = tokens[j];
      if (tok.day < 1 || tok.day > 31) continue;
      const valStart = tok.end;
      const valEnd = (j + 1 < tokens.length) ? tokens[j + 1].start : line.length;
      const value = line.slice(valStart, valEnd).trim();
      if (value) {
        entries.push({ day: tok.day, weekday: tok.weekday, value, raw });
      } else if (j === tokens.length - 1) {
        pending = { day: tok.day, weekday: tok.weekday, raw };
      }
    }
  }

  // フォールバック: 「\d+日」が一切見つからなかった場合の救済
  // 「すみません29も休みに変更で…」のような「日」省略パターン
  if (entries.length === 0) {
    for (let i = 1; i < lines.length; i++) {
      const raw = lines[i];
      const line = toHalfDigits(raw);
      const m = line.match(/(\d{1,2})(?:も|は|を|に|が|から|の)\s*(.+)/);
      if (m) {
        const day = parseInt(m[1], 10);
        if (day >= 1 && day <= 31) {
          entries.push({ day, weekday: null, value: m[2].trim(), raw });
        }
      }
    }
  }

  // 末尾メタ指定「全て○○」を検出 → 種別未指定エントリに type を補完
  // 例: 「全て完全退勤でお願いします」「全部上がり」「全て受付で」
  // 仕様: 既に type が入っているエントリ・OFF系エントリには触らない（保守的）
  if (entries.length > 0) {
    let globalType = '';
    const fullText = String(text || '');
    if (/全(?:て|部)\S{0,15}?(?:完全退勤|退勤|上がり|あがり|上)/.test(fullText)) {
      globalType = '上';
    } else if (/全(?:て|部)\S{0,15}?(?:受け付|受付|受け|受)/.test(fullText)) {
      globalType = '受';
    }
    if (globalType) {
      const hasTypeRe = /受|上|あがり|退勤/;
      const isOffRe = /[❌×✗✖]|休み|やすみ|^休$|^off$/i;
      for (const e of entries) {
        const v = String(e.value || '').trim();
        if (!v) continue;
        if (isOffRe.test(v)) continue;
        if (hasTypeRe.test(v)) continue;
        e.value = v + globalType;
      }
    }
  }

  if (entries.length === 0) return null;
  return { name, entries };
}

/**
 * シフト値の表記ゆれをシート形式に寄せる
 *  - "休み"/"お休み"/"off" → "OFF"
 *  - "当欠"/"前欠"/"当欠店"/"前欠店" → そのまま（文中含有でも検出）
 *  - "19~3受け" / "19〜3受" / "19-3受" → "19-3受"（半角ハイフン統一）
 *  - "19時から23時退勤" → "19-23上"（退勤=上がり=上）
 *  - "21時半から27時退勤" → "21:30-27上"
 *  - "13:00〜22:50上" → "13-22:50上"（:00は省略）
 *  - "12-28上がりでお願いします" → "12-28上"（余計な文字は除去）
 *  - 解析不能なテキスト → ''（スキップ対象）
 */
// "1530" / "930" / "1200" などの軍用時刻を HH:MM に変換
// "12" / "9" / "12:30" はそのまま（先頭ゼロは除去: "09" → "9"）
function parseTimePart(t) {
  t = String(t || '').trim();
  if (/^\d{1,2}$/.test(t)) return String(parseInt(t, 10));
  if (/^\d{1,2}:\d{1,2}$/.test(t)) {
    const [hh, mm] = t.split(':');
    return `${parseInt(hh, 10)}:${mm.padStart(2, '0')}`;
  }
  if (/^\d{3,4}$/.test(t)) {
    const len = t.length;
    const hh = len === 3 ? t.substring(0, 1) : t.substring(0, 2);
    const mm = t.substring(len - 2);
    return `${parseInt(hh, 10)}:${mm}`;
  }
  return t;
}

function normalizeShiftValue(v) {
  // 全角→半角＋空白除去
  let s = toHalfDigits(String(v || '')).replace(/[\s　]+/g, '');
  if (!s) return '';

  // ── v5.2追加: 状況依存OFF（前欠/当欠/店欠の判定マーカー）
  // caller(reflectShiftMessage)で既存セル＋JST投稿日と対象日を比較して最終値を決定する
  if (/やっぱり(?:休み|お休み|おやすみ|休)/.test(s)) return '__YAPPARI_OFF__';
  if (/今日(?:休み|お休み|おやすみ)/.test(s))       return '__TODAY_OFF__';
  if (/店欠|店都合(?:休み)?/.test(s))                return '__TENKETSU__';

  // ── ✗系記号 → OFF（休み）
  if (/[❌×✗✖]/.test(s)) return 'OFF';

  // ── 「休み」「お休み」「おやすみ」「休」「OFF」 → OFF
  if (/休み|やすみ|休$|^休/.test(s) || /^off$/i.test(s)) return 'OFF';

  // ── 欠勤系（文中含有で検出）
  // ※「店欠」は v5.2 で __TENKETSU__ マーカー化されたため、ここでは対象外
  for (const kw of ['当欠店', '前欠店', '当欠', '前欠']) {
    if (s.includes(kw)) return kw;
  }

  // ── 種別検出（受 / 上）
  // ※「あがり」（ひらがな）も検出対象に追加
  let type = '';
  if (/受(?:け|付|け付)?/.test(s)) type = '受';
  else if (/上がり|あがり|退勤|上/.test(s)) type = '上';

  // ── 「ラスト」を時刻に置換: 常に 27（受でも 27、種別不問）
  // ※すいさん仕様「ラストは27時で28時半は使わない」
  // v5.6: 「Ｌ」(全角)「L」(半角・前後英字なしのとき) も 27 として扱う
  s = s.replace(/ラスト|最後|閉店|Ｌ/g, '27');
  s = s.replace(/(?<![A-Za-z])L(?![A-Za-z])/g, '27');

  // ── v5.6: 「19.5」のような小数1桁(.5) → 「:30」に変換（時刻半端表記）
  // 「19.30」より先に処理しないと「19:5」(=19時5分)に誤変換される
  s = s.replace(/(\d{1,2})\.5(?!\d)/g, '$1:30');

  // ── 「19.30」「9.00」のドット表記を「19:30」「9:00」に変換
  // 数字に挟まれたドットのみ（小数点ではなく時刻区切りとして扱う）
  s = s.replace(/(\d{1,2})\.(\d{1,2})/g, '$1:$2');

  // ── 時間/分の漢字を変換 + 範囲記号を半角ハイフンに統一
  let work = s
    .replace(/から/g, '-')
    .replace(/[〜～~ー]/g, '-')
    .replace(/(\d+)時(\d+)分/g, '$1:$2')
    .replace(/(\d+)時半/g, '$1:30')
    .replace(/時/g, '')
    .replace(/分/g, '')
    // 種別キーワードを除去（時刻抽出の邪魔にしない）
    .replace(/受け付|受付|受け|受/g, '')
    .replace(/上がり|あがり|退勤|上/g, '');

  // ── 時刻パターンを抽出（最初に見つかった範囲 or 単発時刻）
  // それ以外の文字（送迎希望／お願いします／🙇／泊まり／に変更／カッコ／etc）は全部無視
  const timeTok = '(?:\\d{3,4}|\\d{1,2}(?::\\d{1,2})?)';
  const rangeRe = new RegExp(`(${timeTok})\\s*-\\s*(${timeTok})`);
  const m = work.match(rangeRe);
  if (m) {
    const simp = (t) => t.replace(/:00$/, '');
    const body = simp(parseTimePart(m[1])) + '-' + simp(parseTimePart(m[2]));
    return type ? body + type : body;
  }

  const singleRe = new RegExp(`(${timeTok})`);
  const single = work.match(singleRe);
  if (single) {
    return parseTimePart(single[1]).replace(/:00$/, '') + (type || '');
  }

  if (type) return type;
  return '';
}

/**
 * 部分時刻更新メッセージを解析
 *  - 「出勤時間13時半に変更」「開始時間9:00」 → {mode:'start', time:'13:30', type:''}
 *  - 「上がり時間20時」「退勤時間ラスト」 → {mode:'end', time:'20', type:''}
 *  - type が明示されていれば付与（既存セルの type を上書き）
 *  - 該当しなければ null
 */
function parsePartialUpdate(text) {
  let s = toHalfDigits(String(text || '')).replace(/[\s　]+/g, '');
  if (!s) return null;

  let mode = '';
  if (/出勤時間|開始時間|始業時間/.test(s)) mode = 'start';
  else if (/上がり時間|あがり時間|退勤時間|終了時間|終わり時間|終業時間/.test(s)) mode = 'end';
  if (!mode) return null;

  // type 検出（明示的にあれば override 用）
  // mode==='end' のときは「上がり/退勤」がトリガー語と被るので type 判定しない
  let type = '';
  if (/受(?:け|付|け付)?/.test(s)) type = '受';
  else if (mode === 'start' && /上がり|あがり|退勤/.test(s)) type = '上';

  // ラスト → 常に 27（v5.6: Ｌ/L も対応）
  s = s.replace(/ラスト|最後|閉店|Ｌ/g, '27');
  s = s.replace(/(?<![A-Za-z])L(?![A-Za-z])/g, '27');
  // v5.6: 小数1桁 .5 → :30
  s = s.replace(/(\d{1,2})\.5(?!\d)/g, '$1:30');
  // ドット → コロン
  s = s.replace(/(\d{1,2})\.(\d{1,2})/g, '$1:$2');
  // 漢字時刻 → 数値
  s = s
    .replace(/(\d+)時(\d+)分/g, '$1:$2')
    .replace(/(\d+)時半/g, '$1:30')
    .replace(/時/g, '')
    .replace(/分/g, '');

  // 単発時刻を抽出（最初に見つかったもの）
  const m = s.match(/(\d{1,2}:\d{1,2}|\d{3,4}|\d{1,2})/);
  if (!m) return null;
  const time = parseTimePart(m[1]).replace(/:00$/, '');
  return { mode, time, type };
}

/**
 * 既存セル値と部分更新指示をマージ
 *  - existing="12:30-19上", partial={mode:'start',time:'13:30'} → "13:30-19上"
 *  - existing="12-19", partial={mode:'end',time:'20'} → "12-20"
 *  - existing="12-19", partial={mode:'start',time:'13',type:'受'} → "13-19受"（type上書き）
 *  - existing が空/OFF/欠勤系 → null（マージ不可）
 *  - existing がパースできない → null
 */
function mergeShiftValue(existing, partial) {
  const e = String(existing || '').trim();
  if (!e) return null;
  if (/^OFF$/i.test(e)) return null;
  if (/^(当欠|前欠|店欠|当欠店|前欠店)$/.test(e)) return null;

  // パターン: "12:30-19上" "12-19" "13-25受" "12" "上"
  const m = e.match(/^(\d{1,2}(?::\d{1,2})?)\s*-\s*(\d{1,2}(?::\d{1,2})?)\s*(受|上)?\s*$/);
  let start, end, type;
  if (m) {
    start = m[1]; end = m[2]; type = m[3] || '';
  } else {
    const m2 = e.match(/^(\d{1,2}(?::\d{1,2})?)\s*(受|上)?\s*$/);
    if (m2) { start = m2[1]; end = ''; type = m2[2] || ''; }
    else return null;
  }

  if (partial.mode === 'start') start = partial.time;
  else if (partial.mode === 'end') end = partial.time;

  const finalType = partial.type || type;
  if (start && end) return `${start}-${end}${finalType}`;
  if (start) return `${start}${finalType}`;
  return null;
}

// ── セル色 ─────────────────────────────────────────────

// C-034 (shift_manager.gs) の CLR と一致させる
function hexToRgb(hex) {
  const m = String(hex).replace('#', '').match(/^([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!m) return { red: 1, green: 1, blue: 1 };
  return {
    red: parseInt(m[1], 16) / 255,
    green: parseInt(m[2], 16) / 255,
    blue: parseInt(m[3], 16) / 255,
  };
}

function getShiftCellStyle(value) {
  const v = String(value || '').trim();
  const up = v.toUpperCase();
  if (!v)                return { bg: '#ffffff', font: '#333333', bold: false };
  if (v === '当欠')       return { bg: '#fecaca', font: '#991b1b', bold: true };
  if (v === '前欠')       return { bg: '#fff1f2', font: '#be123c', bold: true };
  if (v === '店欠')       return { bg: '#fed7aa', font: '#9a3412', bold: true }; // オレンジ（店都合）
  if (v === '当欠店')     return { bg: '#fed7aa', font: '#9a3412', bold: true }; // 互換
  if (v === '前欠店')     return { bg: '#fef9c3', font: '#854d0e', bold: true }; // 互換
  if (up === 'OFF')       return { bg: '#f1f5f9', font: '#94a3b8', bold: false };
  if (v.endsWith('上'))   return { bg: '#dbeafe', font: '#1e40af', bold: false };
  if (v.endsWith('受'))   return { bg: '#dcfce7', font: '#166534', bold: false };
  return                         { bg: '#f3e8ff', font: '#7e22ce', bold: false };
}

/**
 * 初回エントリの属する月を推定:
 *  - 曜日が指定されていれば今月 or 翌月で実際に一致する方を採用
 *  - 曜日なしで、day が今日より過去 かつ day<=7 なら翌月プレビュー扱い
 *  - それ以外は今月
 */
function determineInitialMonth(firstEntry, baseDate) {
  const curYear = baseDate.getFullYear();
  const curMonth = baseDate.getMonth() + 1;
  const todayDay = baseDate.getDate();
  const nextYear = curMonth === 12 ? curYear + 1 : curYear;
  const nextMonth = curMonth === 12 ? 1 : curMonth + 1;

  if (firstEntry.weekday) {
    if (weekdayKanji(curYear, curMonth, firstEntry.day) === firstEntry.weekday) {
      return { year: curYear, month: curMonth };
    }
    if (weekdayKanji(nextYear, nextMonth, firstEntry.day) === firstEntry.weekday) {
      return { year: nextYear, month: nextMonth };
    }
    // どちらにも一致しない → 今月として後段でwarningを出す
    return { year: curYear, month: curMonth };
  }
  // 曜日なし：今日より若い日＆1-7日なら翌月プレビュー扱い
  if (firstEntry.day <= 7 && firstEntry.day < todayDay) {
    return { year: nextYear, month: nextMonth };
  }
  return { year: curYear, month: curMonth };
}

/**
 * 各エントリに year/month を付与
 *  - 初回エントリの月を determineInitialMonth で推定
 *  - 以降は日が戻ったら翌月扱い
 */
function resolveDates(entries, baseDate) {
  if (!entries || entries.length === 0) return [];
  const init = determineInitialMonth(entries[0], baseDate);
  let year = init.year;
  let month = init.month;
  let prevDay = 0;
  const out = [];
  for (const e of entries) {
    if (prevDay > 0 && e.day < prevDay) {
      month += 1;
      if (month > 12) { month = 1; year += 1; }
    }
    prevDay = e.day;
    out.push({ ...e, year, month });
  }
  return out;
}

// ── Sheets/Drive ───────────────────────────────────────

async function findShiftSS(year, month) {
  const auth = createAuth();
  const drive = google.drive({ version: 'v3', auth });
  const titlePart = `${year}年${month}月`;
  const res = await drive.files.list({
    q: `name contains '${titlePart}' and name contains 'シフト表' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: 'files(id, name, modifiedTime)',
    orderBy: 'modifiedTime desc',
    pageSize: 5,
  });
  const files = res.data.files || [];
  return files[0] ? { id: files[0].id, name: files[0].name } : null;
}

async function loadRoster(ssId) {
  const auth = createAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: ssId,
    ranges: ['CREA!A3:A500', 'ふわもこ!A3:A500'],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const roster = { CREA: [], 'ふわもこ': [] };
  const names = ['CREA', 'ふわもこ'];
  (res.data.valueRanges || []).forEach((vr, i) => {
    const rows = vr.values || [];
    rows.forEach((r, idx) => {
      const nm = (r[0] || '').toString();
      if (!nm.trim()) return;
      // 集計行（📊🚨📅 等）はroster対象外
      if (/^(📊|🚨|📅)/.test(nm)) return;
      roster[names[i]].push({
        rowIndex: idx + 3,
        name: nm,
        normalized: normalizeName(nm),
      });
    });
  });
  return roster;
}

// スタッフ管理シートから全キャスト（○有無問わず）をロード
async function loadMaster(ssId) {
  const auth = createAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: ssId,
    range: '👥 スタッフ管理!A2:D500',
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const rows = res.data.values || [];
  const master = { CREA: [], 'ふわもこ': [] };
  rows.forEach((r, idx) => {
    const rowNum = idx + 2; // A2 = row 2
    const creaName = (r[0] || '').toString().trim();
    const creaActive = (r[1] || '').toString().trim() === '○';
    const fuwaName = (r[2] || '').toString().trim();
    const fuwaActive = (r[3] || '').toString().trim() === '○';
    if (creaName && !/^(📊|🚨|📅)/.test(creaName)) {
      master.CREA.push({
        row: rowNum, name: creaName, normalized: normalizeName(creaName),
        active: creaActive, activeCol: 'B',
      });
    }
    if (fuwaName && !/^(📊|🚨|📅)/.test(fuwaName)) {
      master['ふわもこ'].push({
        row: rowNum, name: fuwaName, normalized: normalizeName(fuwaName),
        active: fuwaActive, activeCol: 'D',
      });
    }
  });
  return master;
}

function resolveInMaster(inputName, master) {
  const normIn = normalizeName(inputName);
  if (!normIn || normIn.length < 2) return { matches: [] };
  const buckets = [[], [], [], []];
  for (const store of ['CREA', 'ふわもこ']) {
    for (const e of master[store]) {
      if (!e.normalized) continue;
      if (e.normalized === normIn) buckets[0].push({ store, master: e });
      else if (e.normalized.startsWith(normIn)) buckets[1].push({ store, master: e });
      else if (e.normalized.endsWith(normIn)) buckets[2].push({ store, master: e });
      else if (e.normalized.includes(normIn)) buckets[3].push({ store, master: e });
    }
  }
  for (const b of buckets) {
    if (b.length === 1) return b[0];
    if (b.length > 1) return { matches: b };
  }
  return { matches: [] };
}

/**
 * スタッフを「出勤」状態にしてシフト表に行を追加
 * - 👥 スタッフ管理 の該当行（B or D列）に ○ を書く
 * - シフト表タブの集計行（📊…）の直前に行を挿入
 * - 新行に名前と5種の集計列COUNTIF数式を書く
 * - 既存集計行のCOUNTIF範囲を末尾+1に更新
 */
async function activateStaff(ssId, store, masterEntry, ssYear, ssMonth) {
  const auth = createAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // 1) シート構造を取得
  const meta = await sheets.spreadsheets.get({ spreadsheetId: ssId });
  const targetSheet = meta.data.sheets.find(s => s.properties.title === store);
  if (!targetSheet) throw new Error(`シートタブ「${store}」がSS内に見つかりません`);
  const sheetId = targetSheet.properties.sheetId;

  // 2) A列を走査して集計行の位置を特定
  const aRes = await sheets.spreadsheets.values.get({
    spreadsheetId: ssId,
    range: `${store}!A3:A500`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const aVals = aRes.data.values || [];
  let firstSummaryRow = null;
  const summaryRows = [];
  let lastStaffRow = 2; // ヘッダ行（2）を初期値
  for (let i = 0; i < aVals.length; i++) {
    const rowNum = i + 3;
    const v = ((aVals[i] || [])[0] || '').toString();
    if (/^(📊|🚨|📅|🏪)/.test(v)) {
      if (firstSummaryRow === null) firstSummaryRow = rowNum;
      summaryRows.push(rowNum);
    } else if (v.trim()) {
      if (firstSummaryRow === null) lastStaffRow = rowNum;
    }
  }

  // 挿入位置 = 集計行の先頭 or 末尾スタッフ+1
  const insertAt = firstSummaryRow || (lastStaffRow + 1);

  // 3) 行を挿入（前行からの書式継承はOFF＝色が引き継がれないように）
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: ssId,
    requestBody: {
      requests: [{
        insertDimension: {
          range: {
            sheetId,
            dimension: 'ROWS',
            startIndex: insertAt - 1,
            endIndex: insertAt,
          },
          inheritFromBefore: false,
        },
      },
      // 念のため挿入行の書式を明示的にリセット（当月+翌月プレビュー範囲）
      {
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: insertAt - 1,
            endRowIndex: insertAt,
            startColumnIndex: 1,
            endColumnIndex: 1 + daysInMonth(ssYear, ssMonth) + 7,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 1, green: 1, blue: 1 },
              textFormat: { bold: true, foregroundColor: { red: 0, green: 0, blue: 0 } },
              horizontalAlignment: 'CENTER',
              verticalAlignment: 'MIDDLE',
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
        },
      },
      // C-040 v6.1 (2026-04-30): 新規スタッフのA列に店舗別の名簿色を設定
      // CREA = #93c5fd (薄青) / ふわもこ = #f9a8d4 (薄ピンク)
      {
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: insertAt - 1,
            endRowIndex: insertAt,
            startColumnIndex: 0,
            endColumnIndex: 1,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: store === 'CREA'
                ? { red: 147/255, green: 197/255, blue: 253/255 }   // #93c5fd
                : store === 'ふわもこ'
                  ? { red: 249/255, green: 168/255, blue: 212/255 } // #f9a8d4
                  : { red: 1, green: 1, blue: 1 },                  // 未知店舗は白
              textFormat: { bold: true },
              horizontalAlignment: 'CENTER',
              verticalAlignment: 'MIDDLE',
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat.bold,horizontalAlignment,verticalAlignment)',
        },
      }],
    },
  });

  // 挿入により集計行は +1 ずれる
  const newRow = insertAt;
  const newSummaryRows = summaryRows.map(r => r + 1);
  const newLastStaff = newRow;

  // 4) 新行に名前と集計数式を書く
  //    月の日数から集計列位置を動的算出（30日→AM開始、31日→AN開始）
  const daysInSS = daysInMonth(ssYear, ssMonth);
  const endCol = colIndexToLetter(1 + daysInSS);      // 30日→AE, 31日→AF
  const totalStartCol = daysInSS + 9;                  // 30日→39(AM), 31日→40(AN)
  const rng = `B${newRow}:${endCol}${newRow}`;
  //   AM=出勤 / AN=当欠 / AO=前欠 / AP=店欠(合算: 店欠+当欠店+前欠店)
  const staffFormulas = [
    `=COUNTIF(${rng},"*上*")+COUNTIF(${rng},"*受*")`,
    `=COUNTIF(${rng},"当欠")`,
    `=COUNTIF(${rng},"前欠")`,
    `=COUNTIF(${rng},"店欠")+COUNTIF(${rng},"当欠店")+COUNTIF(${rng},"前欠店")`,
  ];
  const staffUpdates = [
    { range: `${store}!A${newRow}`, values: [[masterEntry.name]] },
  ];
  for (let i = 0; i < staffFormulas.length; i++) {
    staffUpdates.push({
      range: `${store}!${colIndexToLetter(totalStartCol + i)}${newRow}`,
      values: [[staffFormulas[i]]],
    });
  }

  // 5) 集計行のCOUNTIF範囲を更新（B〜合計列：日別式 + 合計式すべて）
  //    "{col}3:{col2}{old}" の末尾行番号を newLastStaff に置換
  const summaryUpdates = [];
  if (newSummaryRows.length > 0) {
    const firstSR = newSummaryRows[0];
    const lastSR = newSummaryRows[newSummaryRows.length - 1];
    const totalColL = colIndexToLetter(totalStartCol);
    const rangeAll = `${store}!B${firstSR}:${totalColL}${lastSR}`;
    const allRes = await sheets.spreadsheets.values.get({
      spreadsheetId: ssId, range: rangeAll, valueRenderOption: 'FORMULA',
    });
    const grid = allRes.data.values || [];
    for (let i = 0; i < newSummaryRows.length; i++) {
      const r = newSummaryRows[i];
      const rowFormulas = grid[i] || [];
      for (let j = 0; j < rowFormulas.length; j++) {
        const f = (rowFormulas[j] || '').toString();
        if (!f || f[0] !== '=') continue;
        const newF = f.replace(/([A-Z]+3:[A-Z]+)\d+/g, `$1${newLastStaff}`);
        if (newF !== f) {
          const colLetter = colIndexToLetter(j + 2);
          summaryUpdates.push({ range: `${store}!${colLetter}${r}`, values: [[newF]] });
        }
      }
    }
  }

  // 6) スタッフ管理に○を書く
  const staffMgrUpdate = {
    range: `👥 スタッフ管理!${masterEntry.activeCol}${masterEntry.row}`,
    values: [['○']],
  };

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: ssId,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [...staffUpdates, ...summaryUpdates, staffMgrUpdate],
    },
  });

  return { store, row: newRow, name: masterEntry.name, activated: true };
}

/**
 * 入力名前を名簿と突き合わせ
 * 優先: exact > prefix > suffix > includes
 */
function resolveStaff(inputName, roster) {
  const normIn = normalizeName(inputName);
  if (!normIn || normIn.length < 2) return { matches: [] };

  const buckets = [[], [], [], []]; // exact, prefix, suffix, includes
  for (const store of ['CREA', 'ふわもこ']) {
    for (const e of roster[store]) {
      if (!e.normalized) continue;
      if (e.normalized === normIn) buckets[0].push({ store, row: e.rowIndex, name: e.name });
      else if (e.normalized.startsWith(normIn)) buckets[1].push({ store, row: e.rowIndex, name: e.name });
      else if (e.normalized.endsWith(normIn)) buckets[2].push({ store, row: e.rowIndex, name: e.name });
      else if (e.normalized.includes(normIn)) buckets[3].push({ store, row: e.rowIndex, name: e.name });
    }
  }
  for (const b of buckets) {
    if (b.length === 1) return b[0];
    if (b.length > 1) return { matches: b };
  }
  return { matches: [] };
}

/**
 * SSの月(ssYear/ssMonth)に対して、書き込む日(year/month/day)の列番号(1-based)を返す
 * 見つからなければ null
 */
function computeColumn(ssYear, ssMonth, targetYear, targetMonth, day) {
  const daysInSS = daysInMonth(ssYear, ssMonth);
  if (targetYear === ssYear && targetMonth === ssMonth) {
    if (day < 1 || day > daysInSS) return null;
    return 1 + day; // B = 2 for day 1
  }
  // 翌月プレビュー領域（day 1〜7 のみ）
  let ssNextY = ssYear, ssNextM = ssMonth + 1;
  if (ssNextM > 12) { ssNextM = 1; ssNextY += 1; }
  if (targetYear === ssNextY && targetMonth === ssNextM) {
    if (day < 1 || day > 7) return null;
    return 1 + daysInSS + day; // AF = 2 + 30 = 32 for April SS day 1
  }
  return null;
}

// ── メイン ─────────────────────────────────────────────

/**
 * シフトメッセージを解析→SSに反映
 * @param {string} text  LINEメッセージ本文
 * @returns {Promise<{type:'ignore'|'success'|'error', message?:string, store?:string, staffName?:string, writtenCount?:number, warnings?:string[], successes?:string[]}>}
 */
async function reflectShiftMessage(text) {
  const parsed = parseShiftMessage(text);
  if (!parsed) return { type: 'ignore' };

  const now = getJSTDate();
  const dated = resolveDates(parsed.entries, now);

  // 各エントリの(year,month)に対応するSSを取得してキャッシュ
  const ssCache = new Map(); // key: "YYYY-M" → { id, name } | null
  async function getSS(y, m) {
    const key = `${y}-${m}`;
    if (ssCache.has(key)) return ssCache.get(key);
    const ss = await findShiftSS(y, m);
    ssCache.set(key, ss);
    return ss;
  }

  // 最初のエントリの月のSSで名簿を解決（どのSSでも同じ名簿を想定）
  const first = dated[0];
  let primarySS = await getSS(first.year, first.month);
  // 見つからなければ、その前月SSのプレビュー領域に入れる可能性を探る
  if (!primarySS) {
    const pm = first.month === 1 ? 12 : first.month - 1;
    const py = first.month === 1 ? first.year - 1 : first.year;
    primarySS = await getSS(py, pm);
  }
  if (!primarySS) {
    return { type: 'error', message: `${first.year}年${first.month}月 のシフト表SSが見つかりません` };
  }

  const rosterCache = new Map();
  const masterCache = new Map();
  const ssMonthCache = new Map(); // ssId → {year, month}
  async function getRoster(ssId) {
    if (rosterCache.has(ssId)) return rosterCache.get(ssId);
    const r = await loadRoster(ssId);
    rosterCache.set(ssId, r);
    return r;
  }
  async function getMaster(ssId) {
    if (masterCache.has(ssId)) return masterCache.get(ssId);
    const m = await loadMaster(ssId);
    masterCache.set(ssId, m);
    return m;
  }
  // SSタイトルから年月を抽出
  function getSSMonthFromName(name) {
    const m = String(name || '').match(/(\d{4})年(\d{1,2})月/);
    return m ? { year: parseInt(m[1], 10), month: parseInt(m[2], 10) } : null;
  }

  // 指定SSでスタッフを解決（シフト表にいれば即返却、いなければmasterからactivate）
  async function ensureStaff(ss, inputName) {
    const roster = await getRoster(ss.id);
    const r = resolveStaff(inputName, roster);
    if (!r.matches) return { ...r, activated: false };
    if (r.matches.length > 1) {
      const list = r.matches.map(m => `${m.name}(${m.store})`).join(' / ');
      throw new Error(`「${inputName}」に複数ヒット: ${list}`);
    }
    // シフト表に居ない → 名簿で探す
    const master = await getMaster(ss.id);
    const mHit = resolveInMaster(inputName, master);
    if (mHit.matches) {
      if (mHit.matches.length === 0) throw new Error(`「${inputName}」が名簿（スタッフ管理）に見つかりません`);
      const list = mHit.matches.map(m => `${m.master.name}(${m.store})`).join(' / ');
      throw new Error(`「${inputName}」が名簿で複数ヒット: ${list}`);
    }
    // 既に○が付いているのにシフト表に居ないケース（整合性乱れ）でも、ここで追加して修復
    const ssYM = ssMonthCache.get(ss.id) || getSSMonthFromName(ss.name);
    if (!ssYM) throw new Error(`SSタイトルから年月を特定できません: ${ss.name}`);
    ssMonthCache.set(ss.id, ssYM);
    const activated = await activateStaff(ss.id, mHit.store, mHit.master, ssYM.year, ssYM.month);
    // ロスターキャッシュを更新
    roster[activated.store].push({
      rowIndex: activated.row,
      name: activated.name,
      normalized: normalizeName(activated.name),
    });
    return activated;
  }

  let primaryStaff;
  try {
    primaryStaff = await ensureStaff(primarySS, parsed.name);
  } catch (err) {
    return { type: 'error', message: err.message };
  }

  // 部分更新で既存セル値を読むため auth/sheets を早めに作る
  const auth = createAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // 書き込み計画: SSごとに {range,value,label} 群を作る
  const plans = new Map(); // ssId → { ss, writes }
  const warnings = [];
  const successes = [];

  for (const e of dated) {
    // 曜日チェック（指定があれば）
    if (e.weekday) {
      const ex = weekdayKanji(e.year, e.month, e.day);
      if (e.weekday !== ex) {
        warnings.push(`${e.month}/${e.day} 曜日不一致(入:${e.weekday}/実:${ex})→スキップ`);
        continue;
      }
    }

    // 部分更新（出勤時間/上がり時間など）か否かを先に判定
    const partial = parsePartialUpdate(e.value);
    let baseValue = null;
    if (!partial) {
      baseValue = normalizeShiftValue(e.value);
      if (!baseValue) {
        warnings.push(`${e.month}/${e.day} 値解析不可「${e.value}」→スキップ`);
        continue;
      }
    }

    // 書き込み対象SSを列挙:
    //   ① 当該エントリの月のSS（本体列）
    //   ② 1〜7日ならその前月SSのプレビュー領域（両方存在すれば両方に書く＝ミラー）
    const targets = [];
    const mainSS = await getSS(e.year, e.month);
    if (mainSS) {
      const col = computeColumn(e.year, e.month, e.year, e.month, e.day);
      if (col) targets.push({ ss: mainSS, col });
    }
    if (e.day <= 7) {
      const pm = e.month === 1 ? 12 : e.month - 1;
      const py = e.month === 1 ? e.year - 1 : e.year;
      const prev = await getSS(py, pm);
      if (prev && (!mainSS || prev.id !== mainSS.id)) {
        const col = computeColumn(py, pm, e.year, e.month, e.day);
        if (col) targets.push({ ss: prev, col });
      }
    }

    if (targets.length === 0) {
      warnings.push(`${e.month}/${e.day} 書き込めるSSなし→スキップ`);
      continue;
    }

    for (const t of targets) {
      let staff = primaryStaff;
      if (t.ss.id !== primarySS.id) {
        try {
          staff = await ensureStaff(t.ss, parsed.name);
        } catch (err) {
          // 月またぎミラー先(前月SS)に新人スタッフが未登録なのは通常運用なので silent skip。
          // メインSSの書込はこのループの別iterationで成功している前提。
          if (/名簿（スタッフ管理）に見つかりません/.test(err.message)) {
            continue;
          }
          warnings.push(`${e.month}/${e.day} ${t.ss.name}: ${err.message}→スキップ`);
          continue;
        }
      }
      const range = `${staff.store}!${colIndexToLetter(t.col)}${staff.row}`;

      // 値の決定: 部分更新／状況依存OFF(前欠/当欠/店欠マーカー)／通常 のいずれか
      let value;
      if (partial) {
        let existing = '';
        try {
          const resp = await sheets.spreadsheets.values.get({
            spreadsheetId: t.ss.id,
            range,
          });
          existing = (resp.data.values && resp.data.values[0] && resp.data.values[0][0]) || '';
        } catch (err) {
          warnings.push(`${e.month}/${e.day} ${t.ss.name} 既存セル読取失敗: ${err.message}→スキップ`);
          continue;
        }
        const merged = mergeShiftValue(existing, partial);
        if (!merged) {
          warnings.push(`${e.month}/${e.day} 既存セル値「${existing || '(空)'}」では部分更新不可→スキップ`);
          continue;
        }
        value = merged;
      } else if (baseValue === '__YAPPARI_OFF__' || baseValue === '__TODAY_OFF__' || baseValue === '__TENKETSU__') {
        // v5.2: 状況依存OFF（前欠/当欠/店欠）の解決
        let existing = '';
        try {
          const resp = await sheets.spreadsheets.values.get({
            spreadsheetId: t.ss.id,
            range,
          });
          existing = (resp.data.values && resp.data.values[0] && resp.data.values[0][0]) || '';
        } catch (err) {
          warnings.push(`${e.month}/${e.day} ${t.ss.name} 既存セル読取失敗: ${err.message}→スキップ`);
          continue;
        }
        const ex = String(existing).trim();
        // Q2: 空欄→警告スキップ（予定なし日には前欠/当欠/店欠は使えない）
        if (!ex) {
          warnings.push(`${e.month}/${e.day} 予定なし(空欄)のため反映不可→スキップ`);
          continue;
        }
        // Q3: OFF→スキップ
        if (/^OFF$/i.test(ex)) {
          warnings.push(`${e.month}/${e.day} 既にOFFのためスキップ`);
          continue;
        }
        // 既に欠勤系→スキップ
        if (/^(当欠|前欠|店欠|当欠店|前欠店)$/.test(ex)) {
          warnings.push(`${e.month}/${e.day} 既に「${ex}」のためスキップ`);
          continue;
        }
        // 解決: JST投稿日と対象日を比較
        const today = getJSTDate();
        const cmp = (e.year - today.getFullYear()) * 10000
                  + (e.month - (today.getMonth() + 1)) * 100
                  + (e.day - today.getDate());
        let resolved;
        if (baseValue === '__TENKETSU__') {
          resolved = '店欠';
        } else if (baseValue === '__TODAY_OFF__') {
          if (cmp !== 0) {
            warnings.push(`${e.month}/${e.day} 「今日休み」だが対象日が今日でない→スキップ`);
            continue;
          }
          resolved = '当欠';
        } else { // __YAPPARI_OFF__
          if (cmp > 0) resolved = '前欠';
          else if (cmp === 0) resolved = '当欠';
          else {
            warnings.push(`${e.month}/${e.day} 過去日への「やっぱり休み」→スキップ`);
            continue;
          }
        }
        value = resolved;
      } else {
        value = baseValue;
      }

      if (!plans.has(t.ss.id)) plans.set(t.ss.id, { ss: t.ss, writes: [] });
      plans.get(t.ss.id).writes.push({
        range, value, label: `${e.month}/${e.day}`,
        rowIndex: staff.row,
        colIndex: t.col,
        store: staff.store,
      });
    }
  }

  // 実書き込み（SSごとに1回のbatchUpdate）＋セル色設定
  let writtenCount = 0;
  for (const [ssId, plan] of plans) {
    const data = plan.writes.map(w => ({ range: w.range, values: [[w.value]] }));
    if (data.length === 0) continue;
    // 値書き込み
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: ssId,
      requestBody: { valueInputOption: 'USER_ENTERED', data },
    });
    // C-040 v6.1 (2026-04-30 すい): セル色付けは main_test.py に移管。
    // 「色付き=ベンリー反映済」「白=未反映/窓外」の意味で統一するため、
    // ここではSS値書込のみで色付けを行わない。
    writtenCount += plan.writes.length;
    plan.writes.forEach(w => successes.push(`${w.label} ${w.value}`));
  }

  return {
    type: 'success',
    store: primaryStaff.store,
    staffName: primaryStaff.name,
    activated: !!primaryStaff.activated,
    writtenCount,
    warnings,
    successes,
  };
}

/**
 * reflectShiftMessage() の結果から LINE 返信テキストを整形
 */
/**
 * LINE 返信テキストを整形（成功時は null を返す＝BOT無言）
 * 返信するのは：エラー / 曜日不一致などの警告が1件以上ある場合のみ
 */
function formatReply(result) {
  if (result.type === 'ignore') return null;
  if (result.type === 'error') return `❌ ${result.message}`;
  // success: 警告がある場合のみ通知（正常時は無言）
  if (!result.warnings || result.warnings.length === 0) return null;
  const lines = [];
  const newMark = result.activated ? '🆕 ' : '';
  lines.push(`⚠️ ${newMark}${result.staffName}【${result.store}】${result.writtenCount}件反映 / ${result.warnings.length}件スキップ`);
  for (const w of result.warnings.slice(0, 5)) lines.push(`・${w}`);
  if (result.warnings.length > 5) lines.push(`（他 ${result.warnings.length - 5}件）`);
  return lines.join('\n');
}

module.exports = {
  reflectShiftMessage,
  formatReply,
  // 以下はテスト用エクスポート
  parseShiftMessage,
  normalizeShiftValue,
  parsePartialUpdate,
  mergeShiftValue,
  normalizeName,
  resolveDates,
  computeColumn,
  weekdayKanji,
  daysInMonth,
};
