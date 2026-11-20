'use strict';
/**
 * Shared Excel helpers for all the "Excel tools" (merger, splitter, matcher ...).
 * Everything uses ExcelJS, so there is only one spreadsheet library to maintain.
 */
const ExcelJS = require('exceljs');

/** An error whose message is safe to show to the user (HTTP 400). */
class UserError extends Error {}

const HEADER_SCAN_ROWS = 15;

/* ------------------------------ cell values ------------------------------ */

/** Turn whatever ExcelJS gives for a cell into a plain value (string/number/Date/null). */
function plain(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if ('result' in v) return plain(v.result); // formula
    if (Array.isArray(v.richText)) return plain(v.richText.map((t) => t.text).join(''));
    if (v.text !== undefined) return plain(v.text); // hyperlink
    return null; // error values etc.
  }
  if (typeof v === 'string') {
    const t = v.trim();
    return t === '' ? null : t;
  }
  return v;
}

const isImeiHeader = (h) => String(h).trim().toLowerCase() === 'imei';

function toImei(v) {
  if (typeof v === 'string' && /^\d{10,}$/.test(v)) return Number(v);
  return v;
}

/** Case-insensitive, whitespace-insensitive header lookup. */
const normHeader = (s) => String(s).toLowerCase().replace(/\s+/g, '');
function findHeader(headers, wanted) {
  return headers.find((h) => normHeader(h) === normHeader(wanted)) || null;
}

/* -------------------------------- reading -------------------------------- */

async function loadWorkbook(buffer, fileName) {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buffer);
  } catch (e) {
    throw new UserError(`"${fileName}" is not a readable .xlsx file.`);
  }
  return wb;
}

/**
 * Reads the first sheet that has a header row. The header row is auto-detected:
 * the first row (within the first 15) containing one of `keys` (default: "IMEI"),
 * so junk title rows such as "Report Table" above the header are skipped.
 *
 * @returns {{headers:string[], rows:object[], sheetName:string, headerRow:number}}
 */
async function readTable(buffer, fileName, { keys = ['imei'] } = {}) {
  const wanted = keys.map((k) => k.toLowerCase());
  const wb = await loadWorkbook(buffer, fileName);

  for (const ws of wb.worksheets) {
    const last = Math.min(ws.rowCount, HEADER_SCAN_ROWS);
    let headerRow = 0;
    for (let r = 1; r <= last; r++) {
      const hit = ws.getRow(r).values.some((c) => {
        const p = plain(c);
        return typeof p === 'string' && wanted.includes(p.toLowerCase());
      });
      if (hit) {
        headerRow = r;
        break;
      }
    }
    if (!headerRow) continue;

    const cols = [];
    ws.getRow(headerRow).eachCell({ includeEmpty: false }, (cell, col) => {
      const name = plain(cell.value);
      if (name !== null) cols.push({ col, name: String(name).trim() });
    });

    const rows = [];
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const rec = {};
      let any = false;
      for (const { col, name } of cols) {
        let v = plain(row.getCell(col).value);
        if (isImeiHeader(name)) v = toImei(v);
        if (v !== null) any = true;
        rec[name] = v;
      }
      if (any) rows.push(rec); // skip fully blank rows
    }
    return { headers: cols.map((c) => c.name), rows, sheetName: ws.name, headerRow };
  }
  throw new UserError(
    `"${fileName}": could not find a header row containing ${keys.map((k) => `"${k.toUpperCase()}"`).join(' or ')}.`,
  );
}

/* -------------------------------- writing -------------------------------- */

const widthFor = (h) => {
  const t = String(h).toLowerCase();
  if (t.includes('timestamp')) return 22;
  if (t === 'imei') return 18;
  return Math.max(12, Math.min(String(h).length + 3, 36));
};

/**
 * Adds a sheet with a bold header row and the given records.
 * `records` items are objects keyed by header; `null` inserts a blank row.
 */
function addTableSheet(wb, name, headers, records, { widths } = {}) {
  const ws = wb.addWorksheet(String(name).replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet1');
  ws.columns = headers.map((h, i) => ({ header: h, key: `c${i}`, width: widths ? widths[i] : widthFor(h) }));
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  headers.forEach((h, i) => {
    if (isImeiHeader(h)) ws.getColumn(i + 1).numFmt = '0';
  });
  for (const rec of records) {
    if (rec === null) ws.addRow([]);
    else ws.addRow(headers.map((h) => (rec[h] === undefined || rec[h] === '' ? null : rec[h])));
  }
  return ws;
}

async function toBuffer(wb) {
  return Buffer.from(await wb.xlsx.writeBuffer());
}

module.exports = {
  ExcelJS,
  UserError,
  plain,
  toImei,
  isImeiHeader,
  findHeader,
  loadWorkbook,
  readTable,
  addTableSheet,
  toBuffer,
};
