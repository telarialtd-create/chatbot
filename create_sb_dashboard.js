'use strict';
/**
 * create_sb_dashboard.js
 * SBシートのデータを元に📈分析シートにダッシュボードを作成
 * レイアウト: 各テーブルの横にチャートを配置（見やすいペア表示）
 */

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = process.argv[2] || '1L5a0SeqSckZARYq3rZBVpBwDBPL4GyliEqA-TIDzW7Y';

function createAuthClient() {
  let client_id, client_secret, refresh_token, access_token;
  if (process.env.GOOGLE_CLIENT_ID) {
    client_id = process.env.GOOGLE_CLIENT_ID;
    client_secret = process.env.GOOGLE_CLIENT_SECRET;
    refresh_token = process.env.GOOGLE_REFRESH_TOKEN;
    access_token = process.env.GOOGLE_ACCESS_TOKEN || null;
  } else {
    const oauthKeys = JSON.parse(fs.readFileSync(path.join(process.env.HOME, '.config/gcp-oauth.keys.json')));
    const credentials = JSON.parse(fs.readFileSync(path.join(process.env.HOME, '.config/gdrive-server-credentials.json')));
    client_id = oauthKeys.installed.client_id;
    client_secret = oauthKeys.installed.client_secret;
    refresh_token = credentials.refresh_token;
    access_token = credentials.access_token;
  }
  const oauth2Client = new google.auth.OAuth2(client_id, client_secret, 'http://localhost');
  oauth2Client.setCredentials({ access_token, refresh_token });
  return oauth2Client;
}

function parseMoney(str) {
  if (!str && str !== 0) return 0;
  if (typeof str === 'number') return str;
  return parseInt(String(str).replace(/[¥￥,\s]/g, '')) || 0;
}

function makeBarChart(sheetId, title, domainRange, seriesConfigs, position, stacked) {
  const chart = {
    spec: {
      title,
      basicChart: {
        chartType: 'BAR',
        legendPosition: seriesConfigs.length > 1 ? 'BOTTOM_LEGEND' : 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: '' },
          { position: 'LEFT_AXIS', title: '' },
        ],
        domains: [{ domain: { sourceRange: { sources: [domainRange] } } }],
        series: seriesConfigs.map(s => ({
          series: { sourceRange: { sources: [s.range] } },
          targetAxis: 'BOTTOM_AXIS',
          color: s.color,
          dataLabel: { type: 'DATA', textFormat: { fontSize: 9 }, placement: 'OUTSIDE_END' },
        })),
        headerCount: 1,
      },
    },
    position: { overlayPosition: position },
  };
  if (stacked) chart.spec.basicChart.stackedType = 'STACKED';
  return { addChart: { chart } };
}

async function main(overrideId) {
  const spreadsheetId = overrideId || SPREADSHEET_ID;
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const metaInfo = await sheets.spreadsheets.get({ spreadsheetId });
  const title = metaInfo.data.properties.title || '';
  const yearMatch = title.match(/(\d{4})/);
  const monthMatch = title.match(/[年\-](\d{1,2})月/);
  const year = yearMatch ? parseInt(yearMatch[1]) : 2026;
  const month = monthMatch ? parseInt(monthMatch[1]) : 4;
  const MONTH_LABEL = `${year}年${month}月`;
  console.log(`📅 対象: ${MONTH_LABEL} (${title})`);

  // ── 1. データ取得 ──
  console.log('📥 SBシートからデータ取得中...');

  const dailyRes = await sheets.spreadsheets.values.get({
    spreadsheetId, range: "'SB'!A1:AO3", valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const dailyRows = dailyRes.data.values || [];
  const dateLabels = dailyRows[0] || [];
  const salaryRow = dailyRows[1] || [];
  const sbRow = dailyRows[2] || [];

  const dailyData = [];
  for (let i = 1; i < dateLabels.length; i++) {
    const label = String(dateLabels[i] || '').trim();
    if (!label || !label.match(/^\d+日$/)) continue;
    const salary = parseMoney(salaryRow[i]);
    const sb = parseMoney(sbRow[i]);
    if (salary === 0 && sb === 0) continue;
    dailyData.push({ day: parseInt(label), salary, sb, total: salary + sb });
  }

  const staffRes = await sheets.spreadsheets.values.get({
    spreadsheetId, range: "'SB'!A8:R39", valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const staffRows = staffRes.data.values || [];

  const staffData = [];
  for (let i = 1; i < staffRows.length; i++) {
    const row = staffRows[i];
    const name = String(row[0] || '').trim();
    if (!name || name.includes('SB1000')) continue;
    const sales = parseMoney(row[1]);
    if (sales === 0) continue;
    const ripiRate = typeof row[7] === 'number' ? Math.round(row[7] * 100) : 0;
    staffData.push({
      name, sales,
      salary: parseMoney(row[2]),
      shimeibk: parseMoney(row[3]),
      free: parseMoney(row[4]),
      ripi: parseMoney(row[5]),
      ripiRate,
      honsu: parseMoney(row[8]),
      totalCost: parseMoney(row[13]),
      profit: parseMoney(row[14]),
      allHours: parseMoney(row[16]),
      profitPerH: parseMoney(row[17]),
    });
  }
  staffData.sort((a, b) => b.sales - a.sales);
  const N = staffData.length;
  console.log(`  日別: ${dailyData.length}日 / スタッフ: ${N}名`);

  // ── 2. シート再作成 ──
  const existing = metaInfo.data.sheets.find(s => s.properties.title === '📈 分析');
  const requests = [];
  if (existing) requests.push({ deleteSheet: { sheetId: existing.properties.sheetId } });
  requests.push({ addSheet: { properties: { title: '📈 分析', index: 1, gridProperties: { rowCount: 700, columnCount: 25 } } } });
  const batchRes = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  const SID = batchRes.data.replies.find(r => r.addSheet)?.addSheet.properties.sheetId;

  // ── 3. データ構築 ──
  // チャートサイズ
  const CW = 550;                      // チャート幅
  const CH = N * 32 + 100;             // チャート高さ（スタッフ用）
  const CHART_ROWS = Math.ceil(CH / 18); // チャートが占める行数
  const CCOL = 6;                       // チャート開始列

  const totalSales = staffData.reduce((s, d) => s + d.sales, 0);
  const totalProfit = staffData.reduce((s, d) => s + d.profit, 0);
  const totalHonsu = staffData.reduce((s, d) => s + d.honsu, 0);
  const totalCost = staffData.reduce((s, d) => s + d.totalCost, 0);
  const avgUnit = totalHonsu > 0 ? Math.round(totalSales / totalHonsu) : 0;
  const profitRate = totalSales > 0 ? Math.round(totalProfit / totalSales * 100) : 0;

  // 各セクションのソート済みデータ
  const byRipi = [...staffData].sort((a, b) => b.ripiRate - a.ripiRate);
  const byHours = [...staffData].sort((a, b) => b.allHours - a.allHours);
  const byProfitH = [...staffData].sort((a, b) => b.profitPerH - a.profitPerH);
  const byHonsu = [...staffData].sort((a, b) => b.honsu - a.honsu);

  // ── 総合評価スコア計算 ──
  // 偏差値方式: 各指標を平均50・標準偏差10に変換して重み付け合算
  function calcDeviation(values) {
    const n = values.length;
    if (n === 0) return [];
    const mean = values.reduce((a, b) => a + b, 0) / n;
    const variance = values.reduce((a, b) => a + (b - mean) ** 2, 0) / n;
    const sd = Math.sqrt(variance) || 1; // 0除算防止
    return values.map(v => 50 + 10 * (v - mean) / sd);
  }

  // 4指標の偏差値を計算
  const profitHValues = staffData.map(s => s.profitPerH);
  const ripiValues = staffData.map(s => s.ripiRate);
  const honsuValues = staffData.map(s => s.honsu);
  const profitRateValues = staffData.map(s => s.sales > 0 ? s.profit / s.sales * 100 : 0);

  const hoursValues = staffData.map(s => s.allHours);
  const profitValues = staffData.map(s => s.profit);

  const devProfitH = calcDeviation(profitHValues);
  const devRipi = calcDeviation(ripiValues);
  const devHonsu = calcDeviation(honsuValues);
  const devProfitRate = calcDeviation(profitRateValues);
  const devHours = calcDeviation(hoursValues);
  const devProfit = calcDeviation(profitValues);

  // 重み: 利益額 35%, 出勤時間 20%, リピ率 15%, 利益/h 15%, 本数 15%
  const WEIGHTS = { profit: 0.35, hours: 0.20, ripi: 0.15, profitH: 0.15, honsu: 0.15 };

  for (let i = 0; i < staffData.length; i++) {
    const totalScore = devProfit[i] * WEIGHTS.profit
                     + devHours[i] * WEIGHTS.hours
                     + devRipi[i] * WEIGHTS.ripi
                     + devProfitH[i] * WEIGHTS.profitH
                     + devHonsu[i] * WEIGHTS.honsu;
    staffData[i].score = Math.round(totalScore * 10) / 10;
    staffData[i].devProfit = Math.round(devProfit[i] * 10) / 10;
    staffData[i].devProfitH = Math.round(devProfitH[i] * 10) / 10;
    staffData[i].devRipi = Math.round(devRipi[i] * 10) / 10;
    staffData[i].devHonsu = Math.round(devHonsu[i] * 10) / 10;
    staffData[i].devHours = Math.round(devHours[i] * 10) / 10;
    // ランク判定（SS〜E）
    if (totalScore >= 65) staffData[i].rank = 'SS';
    else if (totalScore >= 60) staffData[i].rank = 'S';
    else if (totalScore >= 56) staffData[i].rank = 'A';
    else if (totalScore >= 52) staffData[i].rank = 'B';
    else if (totalScore >= 48) staffData[i].rank = 'C';
    else if (totalScore >= 44) staffData[i].rank = 'D';
    else staffData[i].rank = 'E';
  }
  const byScore = [...staffData].sort((a, b) => b.score - a.score);

  // 全データをまとめて構築（各セクションの開始行を記録）
  const rows = [];
  const sec = {}; // セクション位置記録

  // --- KPI ---
  rows.push([`📈 ${MONTH_LABEL} SB分析ダッシュボード`]);
  rows.push([]);
  rows.push(['▼ 月間サマリー']);
  rows.push(['総売上', totalSales, '', '総利益', totalProfit]);
  rows.push(['総本数', totalHonsu, '', '平均単価', avgUnit]);
  rows.push(['総コスト', totalCost, '', '利益率', `${profitRate}%`]);

  // --- 総合評価セクション ---
  rows.push([]);
  sec.score = { headerRow: rows.length, start: rows.length + 1 };
  rows.push(['▼ 総合評価ランキング（利益額 35% ＋ 出勤 20% ＋ リピ率 15% ＋ 利益/h 15% ＋ 本数 15%）']);
  rows.push(['名前', 'ランク', 'スコア', '利益偏差値', '出勤偏差値', 'リピ率偏差値', '利益/h偏差値', '本数偏差値', '利益額', '出勤h', 'リピ率', '本数']);
  for (const s of byScore) {
    rows.push([s.name, s.rank, s.score, s.devProfit, s.devHours, s.devRipi, s.devProfitH, s.devHonsu, s.profit, s.allHours, `${s.ripiRate}%`, s.honsu]);
  }
  sec.score.end = rows.length;
  // チャート分余白
  const scoreGap = Math.max(0, Math.ceil((N * 32 + 100) / 18) - N - 2);
  for (let i = 0; i < scoreGap; i++) rows.push([]);

  // --- 散布図データ（利益/h × リピ率） ---
  rows.push([]);
  sec.scatter = { headerRow: rows.length, start: rows.length + 1 };
  rows.push(['▼ 効率×リピ率マップ（右上が理想）']);
  rows.push(['名前', '利益/h', 'リピ率(%)']);
  for (const s of staffData) {
    rows.push([s.name, s.profitPerH, s.ripiRate]);
  }
  sec.scatter.end = rows.length;
  // 散布図チャート用余白
  for (let i = 0; i < 20; i++) rows.push([]);

  // --- セクション1: 売上ランキング + 売上vs利益チャート ---
  const gap = Math.max(0, CHART_ROWS - N - 2); // テーブルよりチャートが長い分の余白
  const addSection = (label, headers, data, key) => {
    rows.push([]); // 空行
    sec[key] = { headerRow: rows.length, start: rows.length + 1 };
    rows.push(label);
    rows.push(headers);
    for (const d of data) rows.push(d);
    sec[key].end = rows.length;
    // チャート分の余白を確保
    for (let i = 0; i < gap; i++) rows.push([]);
  };

  addSection(
    ['▼ 売上ランキング'],
    ['名前', '売上', '利益', '本数', '利益率'],
    staffData.map(s => [s.name, s.sales, s.profit, s.honsu, `${s.sales > 0 ? Math.round(s.profit / s.sales * 100) : 0}%`]),
    'sales'
  );

  addSection(
    ['▼ 本数ランキング'],
    ['名前', '本数', '売上', '利益', '単価'],
    byHonsu.map(s => [s.name, s.honsu, s.sales, s.profit, s.honsu > 0 ? Math.round(s.sales / s.honsu) : 0]),
    'honsu'
  );

  addSection(
    ['▼ 指名・フリー・リピ内訳（リピ率順）'],
    ['名前', '指名', 'フリー', 'リピ', 'リピ率'],
    byRipi.map(s => [s.name, s.shimeibk, s.free, s.ripi, `${s.ripiRate}%`]),
    'ripi'
  );

  addSection(
    ['▼ 総労働時間ランキング'],
    ['名前', '総時間(h)', '売上', '利益', '利益/h'],
    byHours.map(s => [s.name, s.allHours, s.sales, s.profit, s.profitPerH]),
    'hours'
  );

  addSection(
    ['▼ 利益/h ランキング'],
    ['名前', '利益/h', '総時間(h)', '利益', '売上'],
    byProfitH.map(s => [s.name, s.profitPerH, s.allHours, s.profit, s.sales]),
    'profitH'
  );

  // 日別推移
  rows.push([]);
  sec.daily = { headerRow: rows.length, start: rows.length + 1 };
  rows.push(['▼ 日別 給料・SB推移']);
  rows.push(['日', '給料', 'SB', '合計']);
  for (const d of dailyData) rows.push([`${d.day}日`, d.salary, d.sb, d.total]);
  sec.daily.end = rows.length;

  console.log(`✍️ ${rows.length}行書き込み中...`);
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: "'📈 分析'!A1",
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: rows },
  });

  // ── 4. 書式 + グラフ ──
  console.log('🎨 書式・グラフ作成中...');

  // 範囲ヘルパー
  const R = (startRow, endRow, startCol, endCol) => ({
    sheetId: SID, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: startCol, endColumnIndex: endCol,
  });
  const domain = (s) => R(s.headerRow, s.end, 0, 1);
  const series = (s, col) => R(s.headerRow, s.end, col, col + 1);

  const fmt = [];

  // タイトル
  fmt.push({
    repeatCell: { range: R(0, 1, 0, 12), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 14, foregroundColor: { red: 1, green: 1, blue: 1 } },
      backgroundColor: { red: 0.13, green: 0.13, blue: 0.38 },
    }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' },
  });
  fmt.push({ mergeCells: { range: R(0, 1, 0, 12), mergeType: 'MERGE_ALL' } });

  // KPI背景
  fmt.push({
    repeatCell: { range: R(2, 6, 0, 6), cell: { userEnteredFormat: {
      backgroundColor: { red: 0.93, green: 0.95, blue: 1 },
    }}, fields: 'userEnteredFormat(backgroundColor)' },
  });
  fmt.push({
    repeatCell: { range: R(2, 3, 0, 6), cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
      backgroundColor: { red: 0.23, green: 0.47, blue: 0.71 },
    }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' },
  });
  // KPI金額フォーマット
  fmt.push({ repeatCell: { range: R(3, 6, 1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  fmt.push({ repeatCell: { range: R(3, 6, 4, 5), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' } } }, fields: 'userEnteredFormat(numberFormat)' } });

  // 各セクション共通書式
  for (const key of Object.keys(sec)) {
    const s = sec[key];
    // セクションヘッダー（青帯）
    fmt.push({
      repeatCell: { range: R(s.headerRow, s.headerRow + 1, 0, 6), cell: { userEnteredFormat: {
        textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
        backgroundColor: { red: 0.23, green: 0.47, blue: 0.71 },
      }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' },
    });
    // テーブルヘッダー
    fmt.push({
      repeatCell: { range: R(s.start, s.start + 1, 0, 6), cell: { userEnteredFormat: {
        textFormat: { bold: true },
        backgroundColor: { red: 0.85, green: 0.9, blue: 0.97 },
      }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' },
    });
    // 縞模様
    const dataRows = s.end - s.start - 1;
    for (let i = 0; i < dataRows; i++) {
      if (i % 2 === 1) {
        fmt.push({
          repeatCell: { range: R(s.start + 1 + i, s.start + 2 + i, 0, 5), cell: { userEnteredFormat: {
            backgroundColor: { red: 0.95, green: 0.97, blue: 1 },
          }}, fields: 'userEnteredFormat(backgroundColor)' },
        });
      }
    }
  }

  // 金額フォーマット（各セクション）
  // sales: B,C列
  fmt.push({ repeatCell: { range: R(sec.sales.start + 1, sec.sales.end, 1, 3), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  // honsu: C,D,E列
  fmt.push({ repeatCell: { range: R(sec.honsu.start + 1, sec.honsu.end, 2, 5), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  // hours: C,D,E列
  fmt.push({ repeatCell: { range: R(sec.hours.start + 1, sec.hours.end, 2, 5), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  // profitH: B,D,E列
  fmt.push({ repeatCell: { range: R(sec.profitH.start + 1, sec.profitH.end, 1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  fmt.push({ repeatCell: { range: R(sec.profitH.start + 1, sec.profitH.end, 3, 5), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  // daily: B,C,D列
  fmt.push({ repeatCell: { range: R(sec.daily.start + 1, sec.daily.end, 1, 4), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' } } }, fields: 'userEnteredFormat(numberFormat)' } });

  // 条件付き書式（売上テーブルの利益列）
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [R(sec.sales.start + 1, sec.sales.end, 2, 3)], booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '100000' }] }, format: { backgroundColor: { red: 0.85, green: 0.96, blue: 0.85 } } } }, index: 0 } });
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [R(sec.sales.start + 1, sec.sales.end, 2, 3)], booleanRule: { condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '50000' }] }, format: { backgroundColor: { red: 0.98, green: 0.87, blue: 0.87 } } } }, index: 1 } });

  // 列幅
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 100 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 1, endIndex: 10 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } }); // 区切り列

  // フリーズ
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });

  // === 総合評価セクション書式 ===
  // ヘッダー（青帯）
  fmt.push({
    repeatCell: { range: R(sec.score.headerRow, sec.score.headerRow + 1, 0, 12), cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
      backgroundColor: { red: 0.55, green: 0.15, blue: 0.55 },
    }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' },
  });
  // テーブルヘッダー
  fmt.push({
    repeatCell: { range: R(sec.score.start, sec.score.start + 1, 0, 12), cell: { userEnteredFormat: {
      textFormat: { bold: true },
      backgroundColor: { red: 0.92, green: 0.85, blue: 0.95 },
    }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' },
  });
  // ランク列の色分け（条件付き書式）
  const rankRange = R(sec.score.start + 1, sec.score.end, 1, 2);
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [rankRange], booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'SS' }] }, format: { backgroundColor: { red: 1, green: 0.75, blue: 0 }, textFormat: { bold: true } } } }, index: 0 } });
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [rankRange], booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'S' }] }, format: { backgroundColor: { red: 1, green: 0.9, blue: 0.4 }, textFormat: { bold: true } } } }, index: 1 } });
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [rankRange], booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'A' }] }, format: { backgroundColor: { red: 0.7, green: 0.92, blue: 0.7 }, textFormat: { bold: true } } } }, index: 2 } });
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [rankRange], booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'B' }] }, format: { backgroundColor: { red: 0.85, green: 0.92, blue: 1 } } } }, index: 3 } });
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [rankRange], booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'C' }] }, format: { backgroundColor: { red: 1, green: 0.93, blue: 0.8 } } } }, index: 4 } });
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [rankRange], booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'D' }] }, format: { backgroundColor: { red: 1, green: 0.85, blue: 0.85 } } } }, index: 5 } });
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [rankRange], booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'E' }] }, format: { backgroundColor: { red: 0.85, green: 0.8, blue: 0.8 }, textFormat: { italic: true } } } }, index: 6 } });
  // スコア列のカラースケール（低→高で赤→緑）
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [R(sec.score.start + 1, sec.score.end, 2, 3)], gradientRule: {
    minpoint: { color: { red: 1, green: 0.85, blue: 0.85 }, type: 'MIN' },
    midpoint: { color: { red: 1, green: 1, blue: 0.8 }, type: 'PERCENTILE', value: '50' },
    maxpoint: { color: { red: 0.7, green: 0.92, blue: 0.7 }, type: 'MAX' },
  } }, index: 7 } });
  // 偏差値列もカラースケール
  fmt.push({ addConditionalFormatRule: { rule: { ranges: [R(sec.score.start + 1, sec.score.end, 3, 8)], gradientRule: { // 利益〜本数偏差値
    minpoint: { color: { red: 1, green: 0.9, blue: 0.9 }, type: 'MIN' },
    midpoint: { color: { red: 1, green: 1, blue: 1 }, type: 'PERCENTILE', value: '50' },
    maxpoint: { color: { red: 0.8, green: 0.95, blue: 0.8 }, type: 'MAX' },
  } }, index: 8 } });
  // 縞模様
  for (let i = 0; i < byScore.length; i++) {
    if (i % 2 === 1) {
      fmt.push({
        repeatCell: { range: R(sec.score.start + 1 + i, sec.score.start + 2 + i, 0, 12), cell: { userEnteredFormat: {
          backgroundColor: { red: 0.97, green: 0.95, blue: 1 },
        }}, fields: 'userEnteredFormat(backgroundColor)' },
      });
    }
  }

  // 散布図セクション書式
  fmt.push({
    repeatCell: { range: R(sec.scatter.headerRow, sec.scatter.headerRow + 1, 0, 6), cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
      backgroundColor: { red: 0.13, green: 0.5, blue: 0.35 },
    }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' },
  });
  fmt.push({
    repeatCell: { range: R(sec.scatter.start, sec.scatter.start + 1, 0, 3), cell: { userEnteredFormat: {
      textFormat: { bold: true },
      backgroundColor: { red: 0.85, green: 0.95, blue: 0.88 },
    }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' },
  });

  // === チャート（各テーブルの横に配置） ===
  const CCOL_CHART = 11; // 列幅広がったのでチャート開始列を調整
  const pos = (row) => ({ anchorCell: { sheetId: SID, rowIndex: row, columnIndex: CCOL_CHART }, widthPixels: CW, heightPixels: CH });

  // 1. 売上vs利益 → salesテーブルの横
  fmt.push(makeBarChart(SID, '💰 売上 vs 利益',
    domain(sec.sales),
    [
      { range: series(sec.sales, 1), color: { red: 0.23, green: 0.47, blue: 0.71 } },
      { range: series(sec.sales, 2), color: { red: 0.1, green: 0.65, blue: 0.3 } },
    ],
    pos(sec.sales.headerRow),
  ));

  // 2. 本数ランキング → honsuテーブルの横
  fmt.push(makeBarChart(SID, '🔢 本数ランキング',
    domain(sec.honsu),
    [{ range: series(sec.honsu, 1), color: { red: 0.6, green: 0.2, blue: 0.6 } }],
    pos(sec.honsu.headerRow),
  ));

  // 3. 指名/フリー/リピ → ripiテーブルの横
  fmt.push(makeBarChart(SID, '👤 指名・フリー・リピ 内訳',
    domain(sec.ripi),
    [
      { range: series(sec.ripi, 1), color: { red: 0.93, green: 0.46, blue: 0.13 } },
      { range: series(sec.ripi, 2), color: { red: 0.23, green: 0.47, blue: 0.71 } },
      { range: series(sec.ripi, 3), color: { red: 0.1, green: 0.65, blue: 0.3 } },
    ],
    pos(sec.ripi.headerRow), true,
  ));

  // 4. 総労働時間 → hoursテーブルの横
  fmt.push(makeBarChart(SID, '⏱ 総労働時間ランキング',
    domain(sec.hours),
    [{ range: series(sec.hours, 1), color: { red: 0.2, green: 0.6, blue: 0.8 } }],
    pos(sec.hours.headerRow),
  ));

  // 5. 利益/h → profitHテーブルの横
  fmt.push(makeBarChart(SID, '💹 利益/h ランキング',
    domain(sec.profitH),
    [{ range: series(sec.profitH, 1), color: { red: 0.85, green: 0.35, blue: 0.1 } }],
    pos(sec.profitH.headerRow),
  ));

  // 6. 日別推移 → dailyテーブルの横
  fmt.push({
    addChart: {
      chart: {
        spec: {
          title: '📅 日別 給料・SB推移',
          basicChart: {
            chartType: 'COMBO', legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS', title: '' }, { position: 'LEFT_AXIS', title: '' }],
            domains: [{ domain: { sourceRange: { sources: [domain(sec.daily)] } } }],
            series: [
              { series: { sourceRange: { sources: [series(sec.daily, 1)] } }, targetAxis: 'LEFT_AXIS', type: 'COLUMN', color: { red: 0.23, green: 0.47, blue: 0.71 }, dataLabel: { type: 'DATA', textFormat: { fontSize: 8 }, placement: 'OUTSIDE_END' } },
              { series: { sourceRange: { sources: [series(sec.daily, 2)] } }, targetAxis: 'LEFT_AXIS', type: 'COLUMN', color: { red: 0.9, green: 0.45, blue: 0.1 }, dataLabel: { type: 'DATA', textFormat: { fontSize: 8 }, placement: 'OUTSIDE_END' } },
              { series: { sourceRange: { sources: [series(sec.daily, 3)] } }, targetAxis: 'LEFT_AXIS', type: 'LINE', color: { red: 0.1, green: 0.65, blue: 0.3 } },
            ],
            headerCount: 1,
          },
        },
        position: { overlayPosition: {
          anchorCell: { sheetId: SID, rowIndex: sec.daily.headerRow, columnIndex: CCOL_CHART },
          widthPixels: CW + 50, heightPixels: 380,
        }},
      },
    },
  });

  // === 総合スコア横棒チャート ===
  fmt.push(makeBarChart(SID, '🏆 総合評価スコア（偏差値方式）',
    R(sec.score.start, sec.score.end, 0, 1), // 名前
    [
      { range: R(sec.score.start, sec.score.end, 3, 4), color: { red: 0.85, green: 0.2, blue: 0.1 } },   // 利益偏差値（赤）
      { range: R(sec.score.start, sec.score.end, 4, 5), color: { red: 0.2, green: 0.6, blue: 0.8 } },    // 出勤偏差値（水色）
      { range: R(sec.score.start, sec.score.end, 5, 6), color: { red: 0.1, green: 0.65, blue: 0.3 } },   // リピ率偏差値（緑）
      { range: R(sec.score.start, sec.score.end, 6, 7), color: { red: 0.85, green: 0.55, blue: 0.1 } },  // 利益/h偏差値（オレンジ）
      { range: R(sec.score.start, sec.score.end, 7, 8), color: { red: 0.6, green: 0.2, blue: 0.6 } },    // 本数偏差値（紫）
    ],
    pos(sec.score.headerRow), true // stacked
  ));

  // === 散布図（利益/h × リピ率） ===
  fmt.push({
    addChart: {
      chart: {
        spec: {
          title: '🎯 効率×リピ率マップ（右上が理想）',
          basicChart: {
            chartType: 'SCATTER',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: '利益/h (円)' },
              { position: 'LEFT_AXIS', title: 'リピ率 (%)' },
            ],
            domains: [{ domain: { sourceRange: { sources: [R(sec.scatter.start, sec.scatter.end, 1, 2)] } } }],
            series: [{
              series: { sourceRange: { sources: [R(sec.scatter.start, sec.scatter.end, 2, 3)] } },
              targetAxis: 'LEFT_AXIS',
              color: { red: 0.2, green: 0.4, blue: 0.8 },
              pointStyle: { size: 10 },
              dataLabel: { type: 'DATA', textFormat: { fontSize: 9 }, placement: 'ABOVE', customLabelData: { sourceRange: { sources: [R(sec.scatter.start, sec.scatter.end, 0, 1)] } } },
            }],
            headerCount: 1,
          },
        },
        position: { overlayPosition: {
          anchorCell: { sheetId: SID, rowIndex: sec.scatter.headerRow, columnIndex: CCOL_CHART },
          widthPixels: CW + 50, heightPixels: 420,
        }},
      },
    },
  });

  // フィルター
  fmt.push({ setBasicFilter: { filter: { range: R(sec.sales.start, sec.sales.end, 0, 5), sortSpecs: [{ dimensionIndex: 1, sortOrder: 'DESCENDING' }] } } });

  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: fmt } });

  // フィルタービュー
  console.log('📋 フィルタービュー追加中...');
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        { addFilterView: { filter: { title: '売上 降順', range: R(sec.sales.start, sec.sales.end, 0, 5), sortSpecs: [{ dimensionIndex: 1, sortOrder: 'DESCENDING' }] } } },
        { addFilterView: { filter: { title: '利益 降順', range: R(sec.sales.start, sec.sales.end, 0, 5), sortSpecs: [{ dimensionIndex: 2, sortOrder: 'DESCENDING' }] } } },
        { addFilterView: { filter: { title: '本数 降順', range: R(sec.honsu.start, sec.honsu.end, 0, 5), sortSpecs: [{ dimensionIndex: 1, sortOrder: 'DESCENDING' }] } } },
        { addFilterView: { filter: { title: 'リピ率 降順', range: R(sec.ripi.start, sec.ripi.end, 0, 5), sortSpecs: [{ dimensionIndex: 4, sortOrder: 'DESCENDING' }] } } },
      ],
    },
  });

  console.log('');
  console.log('✅ 完了！レイアウト: テーブル(左) + チャート(右) のペア表示');
  console.log('');
  console.log('  🏆 総合評価ランキング ← → 積み上げ偏差値チャート');
  console.log('  🎯 効率×リピ率マップ  ← → 散布図（右上が理想）');
  console.log('  📊 売上ランキング表   ← → 💰 売上vs利益チャート');
  console.log('  📊 本数ランキング表   ← → 🔢 本数チャート');
  console.log('  📊 指名/リピ内訳表    ← → 👤 積み上げチャート');
  console.log('  📊 労働時間表         ← → ⏱ 時間チャート');
  console.log('  📊 利益/h表           ← → 💹 利益/hチャート');
  console.log('  📊 日別推移表         ← → 📅 日別チャート');
}

// モジュールとして使う場合はmainをexport、直接実行の場合はそのまま実行
module.exports = { runDashboard: main };

if (require.main === module) {
  main().catch(err => {
    console.error('エラー:', err.message);
    if (err.response?.data?.error) console.error('API詳細:', JSON.stringify(err.response.data.error, null, 2));
    process.exit(1);
  });
}
