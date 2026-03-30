'use strict';
/**
 * build_dashboard.js
 * 売上スプレッドシートに「📈 分析」シートを作成し、
 * 曜日別集計・グラフ・条件付き書式を追加する
 */

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1Lmm9jek8RX2_5J-AbHpL1kywUPXZdanQMNov3XV59_4';
const SUMMARY_SHEET_GID = 617109056; // 📊 月次サマリー

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

// 金額文字列 → 数値
function parseMoney(str) {
  if (!str) return 0;
  return parseInt(String(str).replace(/[¥￥,\s]/g, '')) || 0;
}

// 数値文字列 → 数値
function parseNum(str) {
  if (!str) return 0;
  return parseFloat(String(str).replace(/[^\d.]/g, '')) || 0;
}

// 日付インデックス(1始まり) → 曜日名
// 2026年3月1日 = 日曜日(0)
const DAY_NAMES = ['日', '月', '火', '水', '木', '金', '土'];
function getDayName(dayNum) {
  const date = new Date(2026, 2, dayNum); // 月は0始まり
  return DAY_NAMES[date.getDay()];
}

async function main() {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // ── 1. 日別データを読み込む ──
  console.log('📥 月次サマリーから日別データ取得中...');
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'📊 月次サマリー'!A24:I53",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const rawRows = res.data.values || [];

  // 朝/昼/夜の本数を売上シートから取得
  const res2 = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'売上'!G5:Z8", // 朝/昼/夜本数・総本数の行
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const timeRows = res2.data.values || [];
  // row0=朝本数, row1=昼本数, row2=夜本数, row3=総本数 (各列が日付)
  const morningCounts = timeRows[0] || [];
  const noonCounts    = timeRows[1] || [];
  const nightCounts   = timeRows[2] || [];

  // ── 2. 日別データを構造化 ──
  const dailyData = [];
  for (const row of rawRows) {
    const label = String(row[0] || '').trim();
    const dayMatch = label.match(/^(\d+)日$/);
    if (!dayMatch) continue;
    const dayNum = parseInt(dayMatch[1]);
    const sales = parseMoney(row[1]);
    if (sales === 0) continue; // データなし日はスキップ

    const colIdx = dayNum - 1; // 0始まりインデックス
    dailyData.push({
      day: dayNum,
      dayName: getDayName(dayNum),
      sales,
      bookings: parseNum(row[2]),
      staff: parseNum(row[3]),
      girlSalary: parseMoney(row[4]),
      expenses: parseMoney(row[5]),
      avgPerPerson: parseNum(row[6]),
      standby: parseNum(row[7]),
      unitPrice: parseMoney(row[8]),
      morning: parseNum(morningCounts[colIdx]),
      noon: parseNum(noonCounts[colIdx]),
      night: parseNum(nightCounts[colIdx]),
    });
  }

  // ── 3. 曜日別集計 ──
  const byDay = {};
  for (const d of DAY_NAMES) {
    byDay[d] = { count: 0, sales: 0, bookings: 0, unitPrice: 0, avgPP: 0, standby: 0 };
  }
  for (const d of dailyData) {
    const b = byDay[d.dayName];
    b.count++;
    b.sales += d.sales;
    b.bookings += d.bookings;
    b.unitPrice += d.unitPrice;
    b.avgPP += d.avgPerPerson;
    b.standby += d.standby;
  }
  const dayOrder = ['月', '火', '水', '木', '金', '土', '日'];
  const dayStats = dayOrder.map(d => ({
    name: d,
    count: byDay[d].count,
    avgSales: byDay[d].count ? Math.round(byDay[d].sales / byDay[d].count) : 0,
    avgBookings: byDay[d].count ? Math.round(byDay[d].bookings / byDay[d].count * 10) / 10 : 0,
    avgUnitPrice: byDay[d].count ? Math.round(byDay[d].unitPrice / byDay[d].count) : 0,
    avgPP: byDay[d].count ? Math.round(byDay[d].avgPP / byDay[d].count * 10) / 10 : 0,
    avgStandby: byDay[d].count ? Math.round(byDay[d].standby / byDay[d].count * 10) / 10 : 0,
  }));

  // ── 4. 既存の分析シートを削除して再作成 ──
  console.log('📋 分析シートを準備中...');
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const existing = meta.data.sheets.find(s => s.properties.title === '📈 分析');
  const requests = [];

  if (existing) {
    requests.push({ deleteSheet: { sheetId: existing.properties.sheetId } });
  }
  requests.push({
    addSheet: {
      properties: {
        title: '📈 分析',
        index: 1,
        gridProperties: { rowCount: 100, columnCount: 20 },
      },
    },
  });

  const batchRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests },
  });

  const newSheetId = batchRes.data.replies.find(r => r.addSheet)?.addSheet.properties.sheetId;
  console.log(`✅ 分析シート作成完了 (sheetId=${newSheetId})`);

  // ── 5. データ書き込み ──
  console.log('✍️ データ書き込み中...');

  // 月間KPIサマリー読み込み
  const kpiRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'📊 月次サマリー'!B6:B11",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const kpiVals = (kpiRes.data.values || []).map(r => r[0] || '');

  // 目標・達成率
  const targetRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'売上'!B8:D8",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const targetRow = (targetRes.data.values || [[]])[0];
  const target = targetRow[0] || '';
  const achieved = targetRow[2] || '';

  const totalSales = parseMoney(kpiVals[0]);
  const avgDailySales = parseMoney(kpiVals[1]);

  const writeData = [
    // KPIカードセクション
    ['📈 3月 分析ダッシュボード'],
    [],
    ['▼ 月間KPI'],
    ['月間総売上', kpiVals[0], '', '目標', target, '達成率', achieved],
    ['日平均売上', kpiVals[1], '', '平均単価', kpiVals[2]],
    ['カード合計', kpiVals[3], '', 'PayPay合計', kpiVals[4]],
    ['支出合計', kpiVals[5]],
    [],
    // 曜日別集計セクション
    ['▼ 曜日別パフォーマンス'],
    ['曜日', '集計日数', '平均売上', '平均本数', '平均単価', '1人平均', '平均待機(h)'],
    ...dayStats.map(d => [
      d.name,
      d.count,
      d.avgSales,
      d.avgBookings,
      d.avgUnitPrice,
      d.avgPP,
      d.avgStandby,
    ]),
    [],
    // 日別推移セクション（グラフ用）
    ['▼ 日別推移（グラフ用データ）'],
    ['日', '曜日', '総売上', '朝本数', '昼本数', '夜本数', '総本数', '出勤人数', '単価', '1人平均', '待機'],
    ...dailyData.map(d => [
      d.day,
      d.dayName,
      d.sales,
      d.morning,
      d.noon,
      d.night,
      d.bookings,
      d.staff,
      d.unitPrice,
      d.avgPerPerson,
      d.standby,
    ]),
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: "'📈 分析'!A1",
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: writeData },
  });

  // ── 6. グラフ・書式設定 ──
  console.log('🎨 グラフ・書式設定中...');

  // データ行の開始位置
  const dayTableStart = 10; // 0-indexed (行11から)
  const dayTableEnd = dayTableStart + dayStats.length;
  const dailyStart = dayTableEnd + 3; // 日別推移ヘッダー行の次
  const dailyDataStart = dailyStart + 1;
  const dailyDataEnd = dailyDataStart + dailyData.length;

  const formatRequests = [
    // タイトルを大きく
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 7 },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, fontSize: 16, foregroundColor: { red: 1, green: 1, blue: 1 } },
            backgroundColor: { red: 0.13, green: 0.13, blue: 0.38 },
          },
        },
        fields: 'userEnteredFormat(textFormat,backgroundColor)',
      },
    },
    // KPIセクションヘッダー
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 7 },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
            backgroundColor: { red: 0.23, green: 0.47, blue: 0.71 },
          },
        },
        fields: 'userEnteredFormat(textFormat,backgroundColor)',
      },
    },
    // 曜日別ヘッダー背景
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: 8, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1 },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
            backgroundColor: { red: 0.23, green: 0.47, blue: 0.71 },
          },
        },
        fields: 'userEnteredFormat(textFormat,backgroundColor)',
      },
    },
    // 曜日別テーブルヘッダー行
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dayTableStart, endRowIndex: dayTableStart + 1, startColumnIndex: 0, endColumnIndex: 7 },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true },
            backgroundColor: { red: 0.85, green: 0.9, blue: 0.97 },
          },
        },
        fields: 'userEnteredFormat(textFormat,backgroundColor)',
      },
    },
    // 曜日別データ行の縞模様
    ...Array.from({ length: dayStats.length }, (_, i) => ({
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dayTableStart + 1 + i, endRowIndex: dayTableStart + 2 + i, startColumnIndex: 0, endColumnIndex: 7 },
        cell: {
          userEnteredFormat: {
            backgroundColor: i % 2 === 0
              ? { red: 1, green: 1, blue: 1 }
              : { red: 0.95, green: 0.97, blue: 1 },
          },
        },
        fields: 'userEnteredFormat(backgroundColor)',
      },
    })),
    // 土曜を青系
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dayTableStart + 6, endRowIndex: dayTableStart + 7, startColumnIndex: 0, endColumnIndex: 1 },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, foregroundColor: { red: 0.1, green: 0.4, blue: 0.8 } },
          },
        },
        fields: 'userEnteredFormat(textFormat)',
      },
    },
    // 日曜を赤系
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dayTableStart + 7, endRowIndex: dayTableStart + 8, startColumnIndex: 0, endColumnIndex: 1 },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, foregroundColor: { red: 0.8, green: 0.1, blue: 0.1 } },
          },
        },
        fields: 'userEnteredFormat(textFormat)',
      },
    },
    // 平均売上列を数値フォーマット（¥）
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dayTableStart + 1, endRowIndex: dayTableEnd + 1, startColumnIndex: 2, endColumnIndex: 3 },
        cell: {
          userEnteredFormat: {
            numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' },
          },
        },
        fields: 'userEnteredFormat(numberFormat)',
      },
    },
    // 単価列を数値フォーマット
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dayTableStart + 1, endRowIndex: dayTableEnd + 1, startColumnIndex: 4, endColumnIndex: 5 },
        cell: {
          userEnteredFormat: {
            numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' },
          },
        },
        fields: 'userEnteredFormat(numberFormat)',
      },
    },
    // 日別推移ヘッダー
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dailyStart - 1, endRowIndex: dailyStart, startColumnIndex: 0, endColumnIndex: 7 },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
            backgroundColor: { red: 0.23, green: 0.47, blue: 0.71 },
          },
        },
        fields: 'userEnteredFormat(textFormat,backgroundColor)',
      },
    },
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dailyStart, endRowIndex: dailyStart + 1, startColumnIndex: 0, endColumnIndex: 11 },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true },
            backgroundColor: { red: 0.85, green: 0.9, blue: 0.97 },
          },
        },
        fields: 'userEnteredFormat(textFormat,backgroundColor)',
      },
    },
    // 日別推移の総売上列を¥フォーマット
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dailyDataStart, endRowIndex: dailyDataEnd, startColumnIndex: 2, endColumnIndex: 3 },
        cell: {
          userEnteredFormat: {
            numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' },
          },
        },
        fields: 'userEnteredFormat(numberFormat)',
      },
    },
    // 単価列
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: dailyDataStart, endRowIndex: dailyDataEnd, startColumnIndex: 8, endColumnIndex: 9 },
        cell: {
          userEnteredFormat: {
            numberFormat: { type: 'CURRENCY', pattern: '¥#,##0' },
          },
        },
        fields: 'userEnteredFormat(numberFormat)',
      },
    },
    // 日別推移の条件付き書式: 売上が日平均以上 → 緑
    {
      addConditionalFormatRule: {
        rule: {
          ranges: [{
            sheetId: newSheetId,
            startRowIndex: dailyDataStart,
            endRowIndex: dailyDataEnd,
            startColumnIndex: 2,
            endColumnIndex: 3,
          }],
          booleanRule: {
            condition: {
              type: 'NUMBER_GREATER_THAN_EQ',
              values: [{ userEnteredValue: String(avgDailySales) }],
            },
            format: {
              backgroundColor: { red: 0.85, green: 0.96, blue: 0.85 },
            },
          },
        },
        index: 0,
      },
    },
    // 売上が日平均未満 → 薄赤
    {
      addConditionalFormatRule: {
        rule: {
          ranges: [{
            sheetId: newSheetId,
            startRowIndex: dailyDataStart,
            endRowIndex: dailyDataEnd,
            startColumnIndex: 2,
            endColumnIndex: 3,
          }],
          booleanRule: {
            condition: {
              type: 'NUMBER_LESS',
              values: [{ userEnteredValue: String(avgDailySales) }],
            },
            format: {
              backgroundColor: { red: 0.98, green: 0.87, blue: 0.87 },
            },
          },
        },
        index: 1,
      },
    },
    // 列幅調整
    { updateDimensionProperties: { range: { sheetId: newSheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 80 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId: newSheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId: newSheetId, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 110 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId: newSheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 7 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } },
    // タイトル行を結合
    {
      mergeCells: {
        range: { sheetId: newSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 7 },
        mergeType: 'MERGE_ALL',
      },
    },

    // ─── グラフ1: 日別売上推移（折れ線） ───
    {
      addChart: {
        chart: {
          spec: {
            title: '日別売上推移',
            basicChart: {
              chartType: 'COMBO',
              legendPosition: 'BOTTOM_LEGEND',
              axis: [
                { position: 'BOTTOM_AXIS', title: '日付' },
                { position: 'LEFT_AXIS', title: '売上（円）' },
              ],
              domains: [{
                domain: {
                  sourceRange: {
                    sources: [{
                      sheetId: newSheetId,
                      startRowIndex: dailyStart,
                      endRowIndex: dailyDataEnd,
                      startColumnIndex: 0,
                      endColumnIndex: 1,
                    }],
                  },
                },
              }],
              series: [
                {
                  series: {
                    sourceRange: {
                      sources: [{
                        sheetId: newSheetId,
                        startRowIndex: dailyStart,
                        endRowIndex: dailyDataEnd,
                        startColumnIndex: 2,
                        endColumnIndex: 3,
                      }],
                    },
                  },
                  targetAxis: 'LEFT_AXIS',
                  type: 'AREA',
                  color: { red: 0.23, green: 0.47, blue: 0.71 },
                },
              ],
              headerCount: 1,
            },
          },
          position: {
            overlayPosition: {
              anchorCell: { sheetId: newSheetId, rowIndex: 2, columnIndex: 8 },
              widthPixels: 600,
              heightPixels: 280,
            },
          },
        },
      },
    },

    // ─── グラフ2: 朝/昼/夜 積み上げ棒グラフ ───
    {
      addChart: {
        chart: {
          spec: {
            title: '時間帯別本数（朝/昼/夜）',
            basicChart: {
              chartType: 'BAR',
              stackedType: 'STACKED',
              legendPosition: 'BOTTOM_LEGEND',
              axis: [
                { position: 'BOTTOM_AXIS', title: '本数' },
                { position: 'LEFT_AXIS', title: '日付' },
              ],
              domains: [{
                domain: {
                  sourceRange: {
                    sources: [{
                      sheetId: newSheetId,
                      startRowIndex: dailyStart,
                      endRowIndex: dailyDataEnd,
                      startColumnIndex: 0,
                      endColumnIndex: 1,
                    }],
                  },
                },
              }],
              series: [
                {
                  series: {
                    sourceRange: {
                      sources: [{
                        sheetId: newSheetId,
                        startRowIndex: dailyStart,
                        endRowIndex: dailyDataEnd,
                        startColumnIndex: 3,
                        endColumnIndex: 4,
                      }],
                    },
                  },
                  targetAxis: 'BOTTOM_AXIS',
                  color: { red: 0.99, green: 0.73, blue: 0.2 },
                },
                {
                  series: {
                    sourceRange: {
                      sources: [{
                        sheetId: newSheetId,
                        startRowIndex: dailyStart,
                        endRowIndex: dailyDataEnd,
                        startColumnIndex: 4,
                        endColumnIndex: 5,
                      }],
                    },
                  },
                  targetAxis: 'BOTTOM_AXIS',
                  color: { red: 0.23, green: 0.47, blue: 0.71 },
                },
                {
                  series: {
                    sourceRange: {
                      sources: [{
                        sheetId: newSheetId,
                        startRowIndex: dailyStart,
                        endRowIndex: dailyDataEnd,
                        startColumnIndex: 5,
                        endColumnIndex: 6,
                      }],
                    },
                  },
                  targetAxis: 'BOTTOM_AXIS',
                  color: { red: 0.18, green: 0.19, blue: 0.57 },
                },
              ],
              headerCount: 1,
            },
          },
          position: {
            overlayPosition: {
              anchorCell: { sheetId: newSheetId, rowIndex: 18, columnIndex: 8 },
              widthPixels: 600,
              heightPixels: 500,
            },
          },
        },
      },
    },

    // ─── グラフ3: 曜日別平均売上棒グラフ ───
    {
      addChart: {
        chart: {
          spec: {
            title: '曜日別 平均売上',
            basicChart: {
              chartType: 'COLUMN',
              legendPosition: 'NO_LEGEND',
              axis: [
                { position: 'BOTTOM_AXIS', title: '曜日' },
                { position: 'LEFT_AXIS', title: '平均売上（円）' },
              ],
              domains: [{
                domain: {
                  sourceRange: {
                    sources: [{
                      sheetId: newSheetId,
                      startRowIndex: dayTableStart,
                      endRowIndex: dayTableEnd + 1,
                      startColumnIndex: 0,
                      endColumnIndex: 1,
                    }],
                  },
                },
              }],
              series: [{
                series: {
                  sourceRange: {
                    sources: [{
                      sheetId: newSheetId,
                      startRowIndex: dayTableStart,
                      endRowIndex: dayTableEnd + 1,
                      startColumnIndex: 2,
                      endColumnIndex: 3,
                    }],
                  },
                },
                targetAxis: 'LEFT_AXIS',
                color: { red: 0.23, green: 0.47, blue: 0.71 },
              }],
              headerCount: 1,
            },
          },
          position: {
            overlayPosition: {
              anchorCell: { sheetId: newSheetId, rowIndex: 2, columnIndex: 14 },
              widthPixels: 350,
              heightPixels: 280,
            },
          },
        },
      },
    },
  ];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: formatRequests },
  });

  // ── 7. 月次サマリーに条件付き書式を追加 ──
  console.log('🎨 月次サマリーに条件付き書式を追加中...');

  // 日別推移の売上列 (B列 = 1) の行24〜52 (0-indexed: 23〜51)
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{
                sheetId: SUMMARY_SHEET_GID,
                startRowIndex: 23,
                endRowIndex: 52,
                startColumnIndex: 1,
                endColumnIndex: 2,
              }],
              booleanRule: {
                condition: {
                  type: 'NUMBER_GREATER_THAN_EQ',
                  values: [{ userEnteredValue: String(avgDailySales) }],
                },
                format: {
                  backgroundColor: { red: 0.82, green: 0.95, blue: 0.82 },
                  textFormat: { bold: true },
                },
              },
            },
            index: 0,
          },
        },
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{
                sheetId: SUMMARY_SHEET_GID,
                startRowIndex: 23,
                endRowIndex: 52,
                startColumnIndex: 1,
                endColumnIndex: 2,
              }],
              booleanRule: {
                condition: {
                  type: 'NUMBER_LESS',
                  values: [{ userEnteredValue: String(avgDailySales) }],
                },
                format: {
                  backgroundColor: { red: 0.98, green: 0.87, blue: 0.87 },
                },
              },
            },
            index: 1,
          },
        },
      ],
    },
  });

  console.log('');
  console.log('✅ 完了！以下が作成・更新されました：');
  console.log('  📈 分析シート（新規）');
  console.log('    ├─ 月間KPIサマリー');
  console.log('    ├─ 曜日別パフォーマンス表');
  console.log('    ├─ 日別推移テーブル（売上が平均以上=緑/以下=赤）');
  console.log('    ├─ グラフ①: 日別売上推移（エリアチャート）');
  console.log('    ├─ グラフ②: 朝/昼/夜 時間帯別本数（積み上げ棒）');
  console.log('    └─ グラフ③: 曜日別平均売上（棒グラフ）');
  console.log('  📊 月次サマリー（更新）');
  console.log('    └─ 日別売上に条件付き書式（平均超=緑/平均未満=赤）');
  console.log('');
  console.log(`  日平均売上: ¥${avgDailySales.toLocaleString()} を基準に色分け`);
}

main().catch(err => {
  console.error('エラー:', err.message);
  if (err.response?.data?.error) {
    console.error('API詳細:', JSON.stringify(err.response.data.error, null, 2));
  }
  process.exit(1);
});
