/**
 * lib/canvas_render.js
 * Sheetsの矩形範囲を canvas で PNG 化（教えて／明細／明細書 共通）
 */
const { google } = require('googleapis');
const { createCanvas } = require('canvas');
const fs = require('fs');
const path = require('path');
const { createAuthClient, toFullWidth, toRgba, TEMP_DIR } = require('./common');

// A1記法の左上セルから 0-indexed の (row, col) オフセットを計算
function rangeStartOffsets(range) {
  const m = String(range).match(/^([A-Z]+)(\d+)/);
  if (!m) return { rowOffset: 0, colOffset: 0 };
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { rowOffset: parseInt(m[2], 10) - 1, colOffset: col - 1 };
}

// シンプル版（結合非対応）：教えて用
async function screenshotCells(spreadsheetId, gid, range) {
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = meta.data.sheets.find(s => s.properties.sheetId === gid);
  if (!sheet) throw new Error(`シートが見つかりません (gid=${gid})`);
  const sheetName = sheet.properties.title;
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [`'${sheetName}'!${range}`],
    includeGridData: true,
  });
  const gridData = res.data.sheets?.[0]?.data?.[0];

  const rows     = gridData?.rowData || [];
  const colMeta  = gridData?.columnMetadata || [];
  const rowMeta  = gridData?.rowMetadata || [];
  const numCols  = rows.reduce((m, r) => Math.max(m, (r.values || []).length), 0);

  const colWidths  = Array.from({ length: numCols }, (_, i) =>
    colMeta[i]?.pixelSize ? Math.max(colMeta[i].pixelSize * 0.85, 28) : 56);
  const rowHeights = rows.map((_, i) =>
    rowMeta[i]?.pixelSize ? Math.max(rowMeta[i].pixelSize * 0.85, 17) : 20);

  const totalW = Math.round(colWidths.reduce((s, w) => s + w, 0));
  const totalH = Math.round(rowHeights.reduce((s, h) => s + h, 0));

  const canvas = createCanvas(totalW, totalH);
  const ctx    = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, totalW, totalH);

  let y = 0;
  for (let ri = 0; ri < rows.length; ri++) {
    const cells = rows[ri].values || [];
    let x = 0;
    for (let ci = 0; ci < numCols; ci++) {
      const cell  = cells[ci] || {};
      const value = toFullWidth(cell.formattedValue ?? '');
      const fmt   = cell.effectiveFormat || {};
      const tf    = fmt.textFormat || {};
      const w = colWidths[ci], h = rowHeights[ri];

      const bg = fmt.backgroundColor;
      const isWhite = !bg || (bg.red===1 && bg.green===1 && bg.blue===1) || (!bg.red && !bg.green && !bg.blue);
      if (!isWhite) {
        ctx.fillStyle = toRgba(bg);
        ctx.fillRect(x, y, w, h);
      }

      ctx.strokeStyle = '#cccccc';
      ctx.lineWidth = 0.5;
      ctx.strokeRect(x + 0.25, y + 0.25, w - 0.5, h - 0.5);

      if (value) {
        const fontSize = tf.fontSize ? Math.round(tf.fontSize * 0.82) : 9;
        const bold     = tf.bold ? 'bold ' : '';
        ctx.font = `${bold}${fontSize}px "NotoJP"`;
        ctx.fillStyle = tf.foregroundColor ? toRgba(tf.foregroundColor, '#000000') : '#000000';

        const ha = fmt.horizontalAlignment;
        let tx;
        if (ha === 'CENTER')     { ctx.textAlign = 'center'; tx = x + w / 2; }
        else if (ha === 'RIGHT') { ctx.textAlign = 'right';  tx = x + w - 2; }
        else                     { ctx.textAlign = 'left';   tx = x + 2; }
        const ty = y + h / 2 + fontSize * 0.36;

        ctx.save();
        ctx.beginPath();
        ctx.rect(x, y, w, h);
        ctx.clip();
        ctx.fillText(value, tx, ty);
        ctx.restore();
      }
      x += w;
    }
    y += rowHeights[ri];
  }

  if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
  const filename   = `sheet_${Date.now()}.png`;
  const outputPath = path.join(TEMP_DIR, filename);
  await new Promise((resolve, reject) => {
    const out    = fs.createWriteStream(outputPath);
    canvas.createPNGStream().pipe(out);
    out.on('finish', resolve);
    out.on('error', reject);
  });
  setTimeout(() => fs.unlink(outputPath, () => {}), 60 * 60 * 1000);
  return filename;
}

// 結合セル対応版：明細／明細書 用
async function screenshotMeisai(spreadsheetId, sheetName, range, prefix = 'meisai', opts = {}) {
  const { compactEmpty = false } = opts;
  const auth = createAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [`'${sheetName}'!${range}`],
    includeGridData: true,
  });

  const sheetData = res.data.sheets?.[0];
  const gridData  = sheetData?.data?.[0];
  const allMerges = sheetData?.merges || [];

  const rows    = gridData?.rowData || [];
  const colMeta = gridData?.columnMetadata || [];
  const rowMeta = gridData?.rowMetadata || [];
  const numCols = rows.reduce((m, r) => Math.max(m, (r.values || []).length), 0);

  const { rowOffset: ROW_OFFSET, colOffset: COL_OFFSET } = rangeStartOffsets(range);

  const mergeMap  = {};
  const mergedSet = new Set();

  for (const m of allMerges) {
    const r0 = m.startRowIndex    - ROW_OFFSET;
    const c0 = m.startColumnIndex - COL_OFFSET;
    const r1 = m.endRowIndex      - ROW_OFFSET;
    const c1 = m.endColumnIndex   - COL_OFFSET;
    if (r1 <= 0 || c1 <= 0 || r0 >= rows.length || c0 >= numCols) continue;
    const sr = Math.max(r0, 0);
    const sc = Math.max(c0, 0);
    mergeMap[`${sr},${sc}`] = { rowSpan: r1 - sr, colSpan: c1 - sc };
    for (let r = sr; r < r1; r++) {
      for (let c = sc; c < c1; c++) {
        if (r === sr && c === sc) continue;
        mergedSet.add(`${r},${c}`);
      }
    }
  }

  const skipRow = new Set();
  if (compactEmpty) {
    for (let ri = 0; ri < rows.length; ri++) {
      const cells = rows[ri].values || [];
      const hasContent = cells.some(c => {
        if (!c) return false;
        if (c.formattedValue !== undefined && c.formattedValue !== '') return true;
        if (c.userEnteredValue?.boolValue !== undefined) return true;
        return false;
      });
      if (hasContent) continue;
      let inMergeBody = false;
      for (let ci = 0; ci < numCols; ci++) {
        if (mergedSet.has(`${ri},${ci}`)) { inMergeBody = true; break; }
        if (mergeMap[`${ri},${ci}`]) { inMergeBody = true; break; }
      }
      if (!inMergeBody) skipRow.add(ri);
    }
  }

  const colWidths  = Array.from({ length: numCols }, (_, i) =>
    colMeta[i]?.pixelSize ? Math.max(colMeta[i].pixelSize * 0.85, 28) : 56);
  const rowHeights = rows.map((_, i) =>
    rowMeta[i]?.pixelSize ? Math.max(rowMeta[i].pixelSize * 0.85, 17) : 20);

  const totalW = Math.round(colWidths.reduce((s, w) => s + w, 0));
  const totalH = Math.round(rows.reduce((s, _, i) => s + (skipRow.has(i) ? 0 : rowHeights[i]), 0));
  const canvas = createCanvas(totalW, totalH);
  const ctx    = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, totalW, totalH);

  let y = 0;
  for (let ri = 0; ri < rows.length; ri++) {
    if (skipRow.has(ri)) continue;
    const cells = rows[ri].values || [];
    let x = 0;
    for (let ci = 0; ci < numCols; ci++) {
      const w = colWidths[ci];
      const h = rowHeights[ri];

      if (mergedSet.has(`${ri},${ci}`)) { x += w; continue; }

      const cell  = cells[ci] || {};
      const fmt   = cell.effectiveFormat || {};
      const tf    = fmt.textFormat || {};

      const mergeInfo = mergeMap[`${ri},${ci}`];
      let cellW = w, cellH = h;
      if (mergeInfo) {
        for (let dc = 1; dc < mergeInfo.colSpan; dc++) cellW += (colWidths[ci + dc] || 0);
        for (let dr = 1; dr < mergeInfo.rowSpan; dr++) cellH += (rowHeights[ri + dr] || 0);
      }

      const bg = fmt.backgroundColor;
      const isWhite = !bg || (bg.red === 1 && bg.green === 1 && bg.blue === 1) || (!bg.red && !bg.green && !bg.blue);
      if (!isWhite) {
        ctx.fillStyle = toRgba(bg);
        ctx.fillRect(x, y, cellW, cellH);
      }

      ctx.strokeStyle = '#cccccc';
      ctx.lineWidth   = 0.5;
      ctx.strokeRect(x + 0.25, y + 0.25, cellW - 0.5, cellH - 0.5);

      const isBoolean = cell.userEnteredValue?.boolValue !== undefined;
      const value     = isBoolean ? '' : toFullWidth(cell.formattedValue ?? '');
      if (value) {
        const fontSize = tf.fontSize ? Math.round(tf.fontSize * 0.82) : 9;
        const bold     = tf.bold ? 'bold ' : '';
        ctx.font       = `${bold}${fontSize}px "NotoJP"`;
        ctx.fillStyle  = tf.foregroundColor ? toRgba(tf.foregroundColor, '#000000') : '#000000';
        const ha = fmt.horizontalAlignment;
        let tx;
        if (ha === 'CENTER')     { ctx.textAlign = 'center'; tx = x + cellW / 2; }
        else if (ha === 'RIGHT') { ctx.textAlign = 'right';  tx = x + cellW - 2; }
        else                     { ctx.textAlign = 'left';   tx = x + 2; }
        const ty = y + cellH / 2 + fontSize * 0.36;
        ctx.save();
        ctx.beginPath();
        ctx.rect(x, y, cellW, cellH);
        ctx.clip();
        ctx.fillText(value, tx, ty);
        ctx.restore();
      }
      x += w;
    }
    y += rowHeights[ri];
  }

  if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
  const filename   = `${prefix}_${Date.now()}.png`;
  const outputPath = path.join(TEMP_DIR, filename);
  await new Promise((resolve, reject) => {
    const out = fs.createWriteStream(outputPath);
    canvas.createPNGStream().pipe(out);
    out.on('finish', resolve);
    out.on('error', reject);
  });
  setTimeout(() => fs.unlink(outputPath, () => {}), 60 * 60 * 1000);
  return filename;
}

module.exports = { screenshotCells, screenshotMeisai, rangeStartOffsets };
