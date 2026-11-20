'use strict';
/**
 * Report Merger
 * Several "Report Table" exports -> one workbook.
 *   Sheet1 : one header row, then each file's data with a blank row between files
 *   Sheet2 : IMEI1, IMEI2, ... (every unique IMEI, its source file and row count)
 */
const { ExcelJS, UserError, readTable, addTableSheet, toBuffer, isImeiHeader } = require('./io');

async function mergeReports(files) {
  if (!files.length) throw new UserError('Please choose at least one .xlsx file.');

  const parsed = [];
  for (const f of files) parsed.push({ name: f.name, ...(await readTable(f.buffer, f.name)) });

  // column order = first file's order; new columns from later files are appended
  const columns = [];
  for (const p of parsed) for (const h of p.headers) if (!columns.includes(h)) columns.push(h);
  const imeiName = columns.find(isImeiHeader);

  const records = [];
  const imeis = [];
  const fileOf = new Map();
  const rowCount = new Map();
  const notes = [];

  parsed.forEach((p) => {
    if (!p.rows.length) {
      notes.push(`${p.name}: no data rows - skipped`);
      return;
    }
    if (records.length) records.push(null); // blank row between files
    for (const rec of p.rows) {
      records.push(rec);
      const id = rec[imeiName];
      if (id === null || id === undefined) continue;
      if (!fileOf.has(id)) {
        fileOf.set(id, p.name);
        imeis.push(id);
      }
      rowCount.set(id, (rowCount.get(id) || 0) + 1);
    }
    notes.push(`${p.name}: ${p.rows.length.toLocaleString('en-US')} rows`);
  });

  const wb = new ExcelJS.Workbook();
  addTableSheet(wb, 'Sheet1', columns, records);
  addTableSheet(
    wb,
    'Sheet2',
    ['Label', 'IMEI', 'Source File', 'Rows'],
    imeis.map((id, i) => ({ Label: `IMEI${i + 1}`, IMEI: id, 'Source File': fileOf.get(id), Rows: rowCount.get(id) })),
    { widths: [10, 20, 40, 10] },
  );

  const dataRows = records.filter(Boolean).length;
  return {
    buffer: await toBuffer(wb),
    filename: 'merged_report.xlsx',
    summary: { stats: { Files: parsed.length, Rows: dataRows, 'Unique IMEIs': imeis.length }, notes },
  };
}

module.exports = { mergeReports };
