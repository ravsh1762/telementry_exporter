'use strict';
/**
 * Day Shift
 * Adds N days to every TIMESTAMP. By default the output is the date only
 * (YYYY-MM-DD) exactly like the original "1Day++" script; tick "keep time"
 * to get YYYY-MM-DD HH:MM:SS instead.
 */
const { ExcelJS, UserError, readTable, addTableSheet, toBuffer, findHeader } = require('./io');
const { parts, pad } = require('./timeUtils');

async function shiftDays(buffer, fileName, { days = 1, keepTime = false }) {
  if (!Number.isInteger(days)) throw new UserError('Days must be a whole number (use a negative number to go back).');

  const table = await readTable(buffer, fileName, { keys: ['timestamp', 'imei'] });
  const tsCol = findHeader(table.headers, 'timestamp');
  if (!tsCol) throw new UserError(`"${fileName}": no TIMESTAMP column found.`);

  let changed = 0;
  let skipped = 0;
  const rows = table.rows.map((row) => {
    const raw = row[tsCol];
    if (raw === null || raw === undefined) return row;
    const p = parts(raw);
    if (!p) {
      skipped++;
      return row;
    }
    const d = new Date(Date.UTC(p.y, p.m - 1, p.d + days));
    let out = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    if (keepTime) out += ` ${pad(p.H)}:${pad(p.M)}:${pad(p.S)}`;
    changed++;
    return { ...row, [tsCol]: out };
  });

  const wb = new ExcelJS.Workbook();
  addTableSheet(wb, 'Output', table.headers, rows);
  return {
    buffer: await toBuffer(wb),
    filename: 'output.xlsx',
    summary: {
      stats: { Rows: rows.length, 'Timestamps shifted': changed, 'Days added': days },
      notes: skipped ? [`${skipped} timestamp(s) could not be read and were left unchanged.`] : [],
    },
  };
}

module.exports = { shiftDays };
