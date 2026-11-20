'use strict';
/** Last IMEI Extractor - keeps only the LAST row of every IMEI (first-seen order). */
const { ExcelJS, UserError, readTable, addTableSheet, toBuffer, findHeader } = require('./io');

async function extractLastImeiRows(buffer, fileName) {
  const table = await readTable(buffer, fileName, { keys: ['imei'] });
  const imeiCol = findHeader(table.headers, 'imei');
  if (!imeiCol) throw new UserError(`"${fileName}": no IMEI column found.`);

  const last = new Map();
  for (const row of table.rows) {
    const imei = String(row[imeiCol] ?? '').trim();
    if (!imei) continue;
    last.set(imei, row); // a later row overwrites the earlier one
  }

  const wb = new ExcelJS.Workbook();
  addTableSheet(wb, 'Last IMEI Rows', table.headers, [...last.values()]);
  return {
    buffer: await toBuffer(wb),
    filename: 'last_imei_rows.xlsx',
    summary: { stats: { 'Total rows': table.rows.length, 'Unique IMEIs': last.size }, notes: [] },
  };
}

module.exports = { extractLastImeiRows };
