'use strict';
/**
 * IMEI Splitter
 * One sheet -> a ZIP of EXCEL_1.xlsx, EXCEL_2.xlsx ... each holding N IMEIs.
 * Original rows keep their formatting. A blank IMEI cell belongs to the IMEI above it.
 */
const JSZip = require('jszip');
const { ExcelJS, UserError, loadWorkbook, toBuffer } = require('./io');

const cellText = (cell) => {
  const t = cell.text;
  return t === undefined || t === null ? '' : String(t).trim();
};

function copyRow(srcRow, dstSheet, dstRowNo, colCount) {
  const dstRow = dstSheet.getRow(dstRowNo);
  if (srcRow.height) dstRow.height = srcRow.height;
  for (let c = 1; c <= colCount; c++) {
    const s = srcRow.getCell(c);
    const d = dstRow.getCell(c);
    d.value = s.value;
    if (s.style) d.style = JSON.parse(JSON.stringify(s.style));
  }
  dstRow.commit();
}

async function splitByImei(buffer, fileName, { imeiCount, headerRow = 1, imeiColumn = 'IMEI' }) {
  if (!imeiCount || imeiCount < 1) throw new UserError('IMEI count must be 1 or more.');

  const wb = await loadWorkbook(buffer, fileName);
  const ws = wb.worksheets[0];
  if (!ws) throw new UserError('No sheet found in the file.');

  const colCount = ws.columnCount;
  const wanted = String(imeiColumn).trim().toLowerCase();

  let imeiCol = null;
  const hdr = ws.getRow(headerRow);
  for (let c = 1; c <= colCount; c++) {
    if (cellText(hdr.getCell(c)).toLowerCase() === wanted) {
      imeiCol = c;
      break;
    }
  }
  if (!imeiCol) throw new UserError(`Column "${imeiColumn}" not found in header row ${headerRow}.`);

  const imeiOrder = [];
  const rowsByFile = new Map();
  let current = null;
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const v = cellText(ws.getRow(r).getCell(imeiCol));
    if (v) current = v;
    if (current === null) continue;
    if (!imeiOrder.includes(current)) imeiOrder.push(current);
    const fileNo = Math.floor(imeiOrder.indexOf(current) / imeiCount) + 1;
    if (!rowsByFile.has(fileNo)) rowsByFile.set(fileNo, []);
    rowsByFile.get(fileNo).push(r);
  }
  if (!rowsByFile.size) throw new UserError('No IMEI data found.');

  const zip = new JSZip();
  for (const [fileNo, rows] of rowsByFile) {
    const out = new ExcelJS.Workbook();
    const ows = out.addWorksheet(ws.name);
    for (let c = 1; c <= colCount; c++) {
      const w = ws.getColumn(c).width;
      if (w) ows.getColumn(c).width = w;
    }
    for (let r = 1; r <= headerRow; r++) copyRow(ws.getRow(r), ows, r, colCount);
    rows.forEach((srcNo, i) => copyRow(ws.getRow(srcNo), ows, headerRow + 1 + i, colCount));
    zip.file(`EXCEL_${fileNo}.xlsx`, await toBuffer(out));
  }

  return {
    buffer: await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
    filename: 'IMEI_split.zip',
    contentType: 'application/zip',
    summary: {
      stats: { 'Unique IMEIs': imeiOrder.length, 'Files created': rowsByFile.size, 'IMEIs per file': imeiCount },
      notes: [],
    },
  };
}

module.exports = { splitByImei };
