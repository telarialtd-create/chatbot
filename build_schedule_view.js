'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1z1FNqOjM62QGg3USYP09icH_wN18crmT2k03NkKVGj4';
const VIEW_SHEET_NAME = 'スケジュール表';

// 9:00 〜 29:00 を 30分刻み → 40スロット
const START_HOUR = 9;
const END_HOUR   = 29;
const SLOT_MIN   = 30;
const TOTAL_SLOTS = (END_HOUR - START_HOUR) * (60 / SLOT_MIN); // 40

// 列定義 (1-indexed, COLUMN()で使用)
const COL_SLOT0 = 6; // F列がスロット0 (9:00)

// 売上シートの行マッピング
const CREA_START_ROW    = 4;   // 売上 行4
const CREA_END_ROW      = 21;  // 売上 行21 (18行)
const FUWAMOKO_START_ROW = 23; // 売上 行23
const FUWAMOKO_END_ROW   = 48; // 売上 行48 (26行)

// 配色 ― 視認性を重視した明暗コントラスト設計
const DARK_BG    = { red: 0.08, green: 0.08, blue: 0.11 }; // 行背景（濃紺黒）
const HEADER_BG  = { red: 0.11, green: 0.11, blue: 0.18 }; // ヘッダー（濃紺）
const SECTION_BG = { red: 0.06, green: 0.06, blue: 0.10 }; // セクション区切り
const WORKING_BG = { red: 0.18, green: 0.55, blue: 0.95 }; // 勤務時間：鮮やかな青
const DAYOFF_BG  = { red: 0.50, green: 0.07, blue: 0.07 }; // 休み：はっきりした深紅
const UNSET_BG   = { red: 0.13, green: 0.13, blue: 0.17 }; // 非勤務スロット（DARK_BGより少し明るく）
const WHITE_TEXT = { red: 1.0,  green: 1.0,  blue: 1.0  };
const DIM_TEXT   = { red: 0.55, green: 0.58, blue: 0.68 }; // スロットヘッダー文字
const CYAN_TEXT  = { red: 0.30, green: 0.88, blue: 1.0  }; // 出勤ステータス
const RED_TEXT   = { red: 1.0,  green: 0.30, blue: 0.30 }; // 休みステータス
const SLOT_TEXT  = { red: 0.45, green: 0.50, blue: 0.65 }; // タイムラベル
const GOLD_TEXT  = { red: 1.0,  green: 0.82, blue: 0.25 }; // セクション名
const BORDER_CLR = { red: 0.22, green: 0.22, blue: 0.30 }; // グリッド線

let _auth = null;
function createAuthClient() {
  if (_auth) return _auth;
  let client_id, client_secret, refresh_token, access_token;
  if (process.env.GOOGLE_CLIENT_ID) {
    client_id     = process.env.GOOGLE_CLIENT_ID;
    client_secret = process.env.GOOGLE_CLIENT_SECRET;
    refresh_token = process.env.GOOGLE_REFRESH_TOKEN;
    access_token  = process.env.GOOGLE_ACCESS_TOKEN || null;
  } else {
    const keys  = JSON.parse(fs.readFileSync(path.join(process.env.HOME, '.config/gcp-oauth.keys.json')));
    const creds = JSON.parse(fs.readFileSync(path.join(process.env.HOME, '.config/gdrive-server-credentials.json')));
    client_id     = keys.installed.client_id;
    client_secret = keys.installed.client_secret;
    refresh_token = creds.refresh_token;
    access_token  = creds.access_token;
  }
  const client = new google.auth.OAuth2(client_id, client_secret, 'http://localhost');
  client.setCredentials({ access_token, refresh_token });
  _auth = client;
  return client;
}

// 時刻ラベル生成: 9→"9:00", 28.5→"28:30"
function toLabel(h) {
  const hh = Math.floor(h);
  const mm = Math.round((h - hh) * 60);
  return `${hh}:${String(mm).padStart(2, '0')}`;
}

// 売上シートの時刻値（"11", "1130", "14上", "2430上"）を
// 小数時間に変換するスプレッドシート数式
// s = 売上の該当セル参照文字列 (例: "売上!CA4")
function startFormula(s) {
  // CA列は数値のみ (例: 11, 1130, 2030)
  return `=IF(OR(${s}="",${s}=0),"",IF(${s}>=100,INT(${s}/100)+MOD(${s},100)/60,${s}))`;
}

function endFormula(s) {
  // CB列は数値 or テキスト+"上" (例: 16, "14上", "2430上")
  // ISNUMBER で分岐: 数値なら直接、テキストなら REGEXEXTRACT で数字部分を抽出
  return (
    `=IF(OR(${s}="",${s}=0),"",` +
    `IFERROR(` +
      `IF(ISNUMBER(${s}),` +
        `IF(${s}>=100,INT(${s}/100)+MOD(${s},100)/60,${s}),` +
        `LET(n,VALUE(REGEXEXTRACT(TO_TEXT(${s}),"[0-9]+")),` +
          `IF(n>=100,INT(n/100)+MOD(n,100)/60,n))` +
      `),""))`
  );
}

function nameFormula(s) {
  return `=${s}`;
}

function statusFormula(caRef, cdRef) {
  return `=IF(${cdRef}="","",IF(${caRef}<>"","出勤","未設定"))`;
}

function labelFormula(r) {
  return (
    `=IF(OR(C${r}="",D${r}=""),"",` +
    `INT(C${r})&":"&TEXT(MOD(C${r},1)*60,"00")&" 〜 "&` +
    `IF(D${r}>=24,"翌"&(INT(D${r})-24)&":"&TEXT(MOD(D${r},1)*60,"00"),` +
    `INT(D${r})&":"&TEXT(MOD(D${r},1)*60,"00")))`
  );
}

// 予約状態数式（各タイムスロットセルに埋め込む）
// BD=予約開始, BF=予約終了 (Google Sheets 時刻小数×24=時間)
//
// 判定ルール:
//   A. スロット開始時に接客+IV進行中 かつ スロット終了以降まで続く → ×
//   B. スロット開始時に接客+IV進行中 かつ スロット内で終わる:
//        残り ≥ 12分 → "" (空き)   残り 5〜11分 → TEL   残り ≤ 4分 → ×
//   C. 予約がスロット内で新規開始する (必ず < 30分しかない) → ×
//   D. 予約なし かつ シフト終了まで 80分未満 → TEL
//   E. それ以外 → "" (空き)
const MIN_COURSE_H = 80 / 60;

function slotFormula(sheetRow) {
  // インターバルは全員10分固定
  const iv = `10`;

  // ※ Google SheetsのLETではrange変数をSUMPRODUCT内で使うと誤動作するため
  //    range参照はすべてインライン展開する
  const AP = `売上!$AP$3:$AP$120`;
  const BD = `売上!$BD$3:$BD$120`;
  const BF = `売上!$BF$3:$BF$120`;
  const nm = `$A${sheetRow}`;

  // BD/BFが9時(9/24≈0.375)未満 = 翌日早朝扱い → +24して24時超え表記に統一
  // ※0.5(正午)を閾値にすると11:59amなども誤補正されるため9時基準とする
  const BDH = `IF(${BD}<(9/24),${BD}*24+24,${BD}*24)`;
  const BFH = `IF(${BF}<(9/24),${BF}*24+24,${BF}*24)`;

  return (
    `=IF(AND($B${sheetRow}="出勤",$C${sheetRow}<=(COLUMN()-${COL_SLOT0})*0.5+${START_HOUR},$D${sheetRow}>(COLUMN()-${COL_SLOT0})*0.5+${START_HOUR}),` +
    `LET(t,(COLUMN()-${COL_SLOT0})*0.5+${START_HOUR},te,t+0.5,iv,${iv},` +
      // A: スロット開始時に接客+IV進行中 かつ スロット終了以降まで続く → ×
      `ca,SUMPRODUCT((${AP}=${nm})*(${BD}>0)*(${BDH}<=t)*(${BFH}+iv/60>=te))>0,` +
      // B×: 進行中 かつ 残り0〜3分で終了 → ×
      `cbx,SUMPRODUCT((${AP}=${nm})*(${BD}>0)*(${BDH}<=t)*(${BFH}+iv/60>=te-3/60)*(${BFH}+iv/60<te))>0,` +
      // BTEL: 進行中 かつ 残り4〜11分で終了 → TEL
      `cbt,SUMPRODUCT((${AP}=${nm})*(${BD}>0)*(${BDH}<=t)*(${BFH}+iv/60>=te-11/60)*(${BFH}+iv/60<te-3/60))>0,` +
      // C: スロット内で新規予約が開始 → ×
      `cc,SUMPRODUCT((${AP}=${nm})*(${BD}>0)*(${BDH}>t)*(${BDH}<te))>0,` +
      // me: スロット開始時に進行中のIVの終了時刻
      `me,SUMPRODUCT(MAX((${AP}=${nm})*(${BD}>0)*(${BDH}<=t)*(${BFH}+iv/60>t)*(${BFH}+iv/60))),` +
      `avail,IF(me>t,me,t),` +
      // le: フリーブロック先頭 = 直前IV終了時刻 or workStart
      `le,MAX($C${sheetRow},SUMPRODUCT(MAX((${AP}=${nm})*(${BD}>0)*(${BFH}+iv/60<=t)*(${BFH}+iv/60)))),` +
      // D: leベース(空きブロック先頭)からのギャップ
      `cd_le,SUMPRODUCT((${AP}=${nm})*(${BD}>0)*(${BDH}>=te)*(${BDH}<le+${MIN_COURSE_H}+iv/60-2/60))>0,` +
      // availベース: IVがスロット内で終わる場合
      `cd_av,AND(me>t,me<te,SUMPRODUCT((${AP}=${nm})*(${BD}>0)*(${BDH}>=te)*(${BDH}<avail+${MIN_COURSE_H}+iv/60-2/60))>0),` +
      `cd,OR(cd_le,cd_av),` +
      // E: 「上」フラグあり かつ 直前が×(予約中) → ×
      `hasU,ISNUMBER(SEARCH("上",TO_TEXT(IFERROR(INDEX(売上!$CB$4:$CB$60,MATCH(${nm},売上!$CD$4:$CD$60,0)),"")))),` +
      `prevBusy,SUMPRODUCT((${AP}=${nm})*(${BD}>0)*(${BDH}<t)*(${BFH}+iv/60>t-0.5))>0,` +
      `ce,AND(hasU,t+${MIN_COURSE_H}>$D${sheetRow},prevBusy),` +
      `IF(OR(ca,cbx,cc,cd,ce),"×",IF(cbt,"TEL",""))` +
    `),"")`
  );
}

// 行データを生成: 売上の nippoRow → スケジュール表の sheetRow
function makeStaffRow(sheetRow, nippoRow) {
  const slot = slotFormula(sheetRow);
  return [
    nameFormula(`売上!CD${nippoRow}`),
    statusFormula(`売上!CA${nippoRow}`, `売上!CD${nippoRow}`),
    startFormula(`売上!CA${nippoRow}`),
    endFormula(`売上!CB${nippoRow}`),
    labelFormula(sheetRow),
    ...new Array(TOTAL_SLOTS).fill(slot),
  ];
}

async function main() {
  const auth   = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // ①シート取得 or 作成
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  let viewSheetId = null;
  const existing = meta.data.sheets.find(s => s.properties.title === VIEW_SHEET_NAME);

  if (existing) {
    viewSheetId = existing.properties.sheetId;
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${VIEW_SHEET_NAME}'`,
    });
    // 既存の条件付き書式を全削除
    const cfSheet = meta.data.sheets.find(s => s.properties.sheetId === viewSheetId);
    const cfCount = cfSheet && cfSheet.conditionalFormats ? cfSheet.conditionalFormats.length : 0;
    if (cfCount > 0) {
      const deleteReqs = [];
      for (let i = cfCount - 1; i >= 0; i--) {
        deleteReqs.push({ deleteConditionalFormatRule: { sheetId: viewSheetId, index: i } });
      }
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { requests: deleteReqs },
      });
    }
    console.log('既存シートをクリア');
  } else {
    const addRes = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          addSheet: {
            properties: {
              title: VIEW_SHEET_NAME,
              gridProperties: {
                rowCount: 60,
                columnCount: TOTAL_SLOTS + COL_SLOT0 + 2,
              },
            },
          },
        }],
      },
    });
    viewSheetId = addRes.data.replies[0].addSheet.properties.sheetId;
    console.log('新しいシートを作成:', viewSheetId);
  }

  // ②データ行を構築
  const slotLabels = Array.from({ length: TOTAL_SLOTS }, (_, i) =>
    toLabel(START_HOUR + i * SLOT_MIN / 60)
  );
  const headerRow = ['名前', '状態', '出勤', '退勤', '時間帯', ...slotLabels];

  // CREA rows (売上 4〜21 → シート行 2〜19)
  const creaRows = [];
  for (let nippoRow = CREA_START_ROW; nippoRow <= CREA_END_ROW; nippoRow++) {
    const sheetRow = nippoRow - CREA_START_ROW + 2;
    creaRows.push(makeStaffRow(sheetRow, nippoRow));
  }

  // セクション区切り (シート行 20)
  // header(1行) + CREA行数(18行) + 1 = 20行目
  const SEP_ROW_INDEX = 1 + (CREA_END_ROW - CREA_START_ROW + 1) + 1; // = 20 (1-indexed)
  const separatorRow = ['▼ ふわもこSPA', '', '', '', '', ...new Array(TOTAL_SLOTS).fill('')];

  // ふわもこ rows (売上 23〜48 → シート行 21〜46)
  const fuwaRows = [];
  for (let nippoRow = FUWAMOKO_START_ROW; nippoRow <= FUWAMOKO_END_ROW; nippoRow++) {
    const sheetRow = nippoRow - FUWAMOKO_START_ROW + SEP_ROW_INDEX + 1;
    fuwaRows.push(makeStaffRow(sheetRow, nippoRow));
  }

  const allRows = [headerRow, ...creaRows, separatorRow, ...fuwaRows];
  const totalDataRows = allRows.length;

  // ③値を書き込み
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${VIEW_SHEET_NAME}'!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: allRows },
  });
  console.log(`値を書き込み完了 (${totalDataRows} 行)`);

  // ④書式リクエストを構築
  const requests = [];
  const totalCols = TOTAL_SLOTS + COL_SLOT0 - 1; // 0-indexed 終端

  // シートプロパティ: フリーズ・タブ色
  requests.push({
    updateSheetProperties: {
      properties: {
        sheetId: viewSheetId,
        gridProperties: { frozenRowCount: 1, frozenColumnCount: 2 },
        tabColor: { red: 0.2, green: 0.5, blue: 0.9 },
      },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount,tabColor',
    },
  });

  // 列幅
  [
    [0, 1, 110],                       // A: 名前
    [1, 2, 68],                        // B: 状態
    [2, 3, 52],                        // C: 出勤
    [3, 4, 52],                        // D: 退勤
    [4, 5, 145],                       // E: 時間帯
    [5, COL_SLOT0 + TOTAL_SLOTS, 28],  // F〜: タイムスロット
  ].forEach(([start, end, px]) => {
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: viewSheetId, dimension: 'COLUMNS', startIndex: start, endIndex: end },
        properties: { pixelSize: px },
        fields: 'pixelSize',
      },
    });
  });

  // 行高（ヘッダー・データ・セクションで差をつける）
  requests.push({
    updateDimensionProperties: {
      range: { sheetId: viewSheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 30 },
      fields: 'pixelSize',
    },
  });
  requests.push({
    updateDimensionProperties: {
      range: { sheetId: viewSheetId, dimension: 'ROWS', startIndex: 1, endIndex: totalDataRows + 1 },
      properties: { pixelSize: 38 },
      fields: 'pixelSize',
    },
  });

  // シート全体のデフォルト背景
  requests.push({
    repeatCell: {
      range: { sheetId: viewSheetId, startRowIndex: 0, endRowIndex: totalDataRows + 1,
               startColumnIndex: 0, endColumnIndex: totalCols + 1 },
      cell: {
        userEnteredFormat: {
          backgroundColor: DARK_BG,
          textFormat: { foregroundColor: DIM_TEXT, fontSize: 9 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    },
  });

  // ヘッダー行
  requests.push({
    repeatCell: {
      range: { sheetId: viewSheetId, startRowIndex: 0, endRowIndex: 1,
               startColumnIndex: 0, endColumnIndex: totalCols + 1 },
      cell: {
        userEnteredFormat: {
          backgroundColor: HEADER_BG,
          textFormat: { foregroundColor: SLOT_TEXT, bold: true, fontSize: 8 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  // ヘッダー: 最初の2列は白テキスト
  requests.push({
    repeatCell: {
      range: { sheetId: viewSheetId, startRowIndex: 0, endRowIndex: 1,
               startColumnIndex: 0, endColumnIndex: 2 },
      cell: {
        userEnteredFormat: {
          textFormat: { foregroundColor: WHITE_TEXT, bold: true, fontSize: 9 },
        },
      },
      fields: 'userEnteredFormat.textFormat',
    },
  });

  // セクション区切り行 (0-indexed: SEP_ROW_INDEX)
  const sepRowIdx = SEP_ROW_INDEX - 1; // 0-indexed
  requests.push({
    repeatCell: {
      range: { sheetId: viewSheetId, startRowIndex: sepRowIdx, endRowIndex: sepRowIdx + 1,
               startColumnIndex: 0, endColumnIndex: totalCols + 1 },
      cell: {
        userEnteredFormat: {
          backgroundColor: SECTION_BG,
          textFormat: { foregroundColor: GOLD_TEXT, bold: true, fontSize: 10 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    },
  });

  // データ行: 名前列（白 + 太字）
  requests.push({
    repeatCell: {
      range: { sheetId: viewSheetId, startRowIndex: 1, endRowIndex: totalDataRows + 1,
               startColumnIndex: 0, endColumnIndex: 1 },
      cell: {
        userEnteredFormat: {
          textFormat: { foregroundColor: WHITE_TEXT, fontSize: 10, bold: true },
        },
      },
      fields: 'userEnteredFormat.textFormat',
    },
  });

  // データ行: 時間帯列（明るめにして読みやすく）
  requests.push({
    repeatCell: {
      range: { sheetId: viewSheetId, startRowIndex: 1, endRowIndex: totalDataRows + 1,
               startColumnIndex: 4, endColumnIndex: 5 },
      cell: {
        userEnteredFormat: {
          textFormat: { foregroundColor: { red: 0.78, green: 0.85, blue: 1.0 }, fontSize: 9, bold: true },
          horizontalAlignment: 'LEFT',
        },
      },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
    },
  });

  // タイムスロット列: デフォルト（非勤務スロット）
  requests.push({
    repeatCell: {
      range: { sheetId: viewSheetId, startRowIndex: 1, endRowIndex: totalDataRows + 1,
               startColumnIndex: COL_SLOT0 - 1, endColumnIndex: totalCols + 1 },
      cell: {
        userEnteredFormat: {
          backgroundColor: UNSET_BG,
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
    },
  });

  // タイムスロット全体にグリッド線（縦・横）を追加
  requests.push({
    updateBorders: {
      range: {
        sheetId: viewSheetId,
        startRowIndex: 0, endRowIndex: totalDataRows + 1,
        startColumnIndex: COL_SLOT0 - 1, endColumnIndex: totalCols + 1,
      },
      innerVertical:   { style: 'SOLID', color: BORDER_CLR, width: 1 },
      innerHorizontal: { style: 'SOLID', color: BORDER_CLR, width: 1 },
      top:    { style: 'SOLID', color: BORDER_CLR, width: 1 },
      bottom: { style: 'SOLID', color: BORDER_CLR, width: 1 },
      left:   { style: 'SOLID', color: BORDER_CLR, width: 1 },
      right:  { style: 'SOLID', color: BORDER_CLR, width: 1 },
    },
  });

  // 1時間ごとの区切り（30分スロット2つで1時間）に太めのボーダーを追加
  for (let hour = 0; hour <= (END_HOUR - START_HOUR); hour++) {
    const colIdx = COL_SLOT0 - 1 + hour * 2; // 各時間の左端列（0-indexed）
    if (colIdx > totalCols) break;
    requests.push({
      updateBorders: {
        range: {
          sheetId: viewSheetId,
          startRowIndex: 0, endRowIndex: totalDataRows + 1,
          startColumnIndex: colIdx, endColumnIndex: colIdx + 1,
        },
        left: { style: 'SOLID', color: { red: 0.35, green: 0.38, blue: 0.52 }, width: 2 },
      },
    });
  }

  // 左側（名前〜時間帯列）の右端に区切り線
  requests.push({
    updateBorders: {
      range: {
        sheetId: viewSheetId,
        startRowIndex: 0, endRowIndex: totalDataRows + 1,
        startColumnIndex: COL_SLOT0 - 2, endColumnIndex: COL_SLOT0 - 1,
      },
      right: { style: 'SOLID', color: { red: 0.35, green: 0.38, blue: 0.52 }, width: 2 },
    },
  });

  // ⑤条件付き書式（優先順位: index 小さい方が高い）

  // ルール0a: 【最優先】× セル → 接客中/インターバル → 深いオレンジ
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [{
          sheetId: viewSheetId,
          startRowIndex: 1, endRowIndex: totalDataRows + 1,
          startColumnIndex: COL_SLOT0 - 1,
          endColumnIndex: COL_SLOT0 - 1 + TOTAL_SLOTS,
        }],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: '×' }] },
          format: {
            backgroundColor: { red: 0.85, green: 0.25, blue: 0.10 },
            textFormat: { foregroundColor: WHITE_TEXT, bold: true },
          },
        },
      },
      index: 0,
    },
  });

  // ルール0b: TEL セル → 黄色
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [{
          sheetId: viewSheetId,
          startRowIndex: 1, endRowIndex: totalDataRows + 1,
          startColumnIndex: COL_SLOT0 - 1,
          endColumnIndex: COL_SLOT0 - 1 + TOTAL_SLOTS,
        }],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'TEL' }] },
          format: {
            backgroundColor: { red: 0.95, green: 0.78, blue: 0.08 },
            textFormat: { foregroundColor: { red: 0.10, green: 0.10, blue: 0.10 }, bold: true },
          },
        },
      },
      index: 1,
    },
  });

  // ルール1: 出勤 → 勤務時間帯を水色
  // slot_time = (COLUMN() - COL_SLOT0) * 0.5 + START_HOUR
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [{
          sheetId: viewSheetId,
          startRowIndex: 1, endRowIndex: totalDataRows + 1,
          startColumnIndex: COL_SLOT0 - 1,
          endColumnIndex: COL_SLOT0 - 1 + TOTAL_SLOTS,
        }],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{
              userEnteredValue:
                `=AND($B2="出勤",$C2<=(COLUMN()-${COL_SLOT0})*0.5+${START_HOUR},$D2>(COLUMN()-${COL_SLOT0})*0.5+${START_HOUR})`,
            }],
          },
          format: { backgroundColor: WORKING_BG },
        },
      },
      index: 2,
    },
  });

  // ルール3: 休み → スロット列を赤
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [{
          sheetId: viewSheetId,
          startRowIndex: 1, endRowIndex: totalDataRows + 1,
          startColumnIndex: COL_SLOT0 - 1,
          endColumnIndex: COL_SLOT0 - 1 + TOTAL_SLOTS,
        }],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: `=$B2="休み"` }],
          },
          format: { backgroundColor: DAYOFF_BG },
        },
      },
      index: 3,
    },
  });

  // ルール4: 休み → 状態列を赤テキスト
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [{
          sheetId: viewSheetId,
          startRowIndex: 1, endRowIndex: totalDataRows + 1,
          startColumnIndex: 1, endColumnIndex: 2,
        }],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: `=$B2="休み"` }],
          },
          format: { textFormat: { foregroundColor: RED_TEXT, bold: true } },
        },
      },
      index: 4,
    },
  });

  // ルール5: 出勤 → 状態列をシアンテキスト
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [{
          sheetId: viewSheetId,
          startRowIndex: 1, endRowIndex: totalDataRows + 1,
          startColumnIndex: 1, endColumnIndex: 2,
        }],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: `=$B2="出勤"` }],
          },
          format: { textFormat: { foregroundColor: CYAN_TEXT, bold: true } },
        },
      },
      index: 5,
    },
  });

  // ⑥一括送信 (500件単位)
  const CHUNK = 500;
  for (let i = 0; i < requests.length; i += CHUNK) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: requests.slice(i, i + CHUNK) },
    });
  }
  console.log(`書式設定完了 (${requests.length} リクエスト)`);

  console.log(`\n✅ 完成: https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`);
}

main().catch(err => {
  console.error('エラー:', err.message || err);
  process.exit(1);
});
