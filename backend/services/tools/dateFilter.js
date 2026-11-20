'use strict';
/**
 * Date Filter
 * Keeps only the rows whose TIMESTAMP (or DATE) is inside From..To (inclusive),
 * grouped by IMEI with a blank row between IMEIs.
 */
const { ExcelJS, UserError, readTable, addTableSheet, toBuffer, findHeader } = require('./io');
const { dateKey } = require('./timeUtils');

async function filterByDate(buffer, fileName, { fromDate, toDate }) {
  const from = dateKey(fromDate);
  const to = dateKey(toDate || fromDate);
  if (!from) throw new UserError('Please give a valid From date (YYYY-MM-DD).');
  if (!to) throw new UserError('Please give a valid To date (YYYY-MM-DD).');
  if (from > to) throw new UserError('To date must be on or after From date.');

  const table = await readTable(buffer, fileName, { keys: ['timestamp', 'imei'] });
  const tsCols = ['timestamp', 'date'].map((n) => findHeader(table.headers, n)).filter(Boolean);
  if (!tsCols.length) throw new UserError(`"${fileName}": no TIMESTAMP (or DATE) column found.`);
  const imeiCol = findHeader(table.headers, 'imei');

  const groups = new Map();
  let matched = 0;
  for (const row of table.rows) {
    const raw = tsCols.map((c) => row[c]).find((v) => v !== null && v !== undefined && v !== '');
    const d = raw === undefined ? null : dateKey(raw);
    if (!d || d < from || d > to) continue;
    const imei = (imeiCol && row[imeiCol]) || 'UNKNOWN';
    if (!groups.has(imei)) groups.set(imei, []);
    groups.get(imei).push(row);
    matched++;
  }

  const records = [];
  for (const rows of groups.values()) {
    if (records.length) records.push(null); // blank row between IMEIs
    records.push(...rows);
  }

  const wb = new ExcelJS.Workbook();
  addTableSheet(wb, `${from}_to_${to}`, table.headers, records);
  return {
    buffer: await toBuffer(wb),
    filename: `${from}_to_${to}.xlsx`,
    summary: {
      stats: { 'Rows in file': table.rows.length, 'Rows matched': matched, IMEIs: groups.size },
      notes: [`Date range: ${from} to ${to}`],
    },
  };
}

module.exports = { filterByDate };
