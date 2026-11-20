'use strict';
/**
 * Timestamp Matcher
 * Matches Excel 1 rows against Excel 2 rows by IMEI + exact timestamp.
 * Output workbook (first row + first column frozen, borders everywhere):
 *   Matches        - same timestamp found in both files
 *   Missing in Ex2 - timestamps present in Excel 1 but not in Excel 2
 *   My Color Sheet - date-wise IMEI status (green = all matched, red = something missing)
 */
const { ExcelJS, UserError, readTable, addTableSheet, toBuffer, findHeader } = require('./io');
const { pad, tsKey, limitKey, show } = require('./timeUtils');

const DEFAULTS = {
  ex1: { imei: 'IMEI', timestamp: 'TIMESTAMP', value: 'POPTOTWD1' },
  ex2: { imei: 'IMEI', timestamp: 'timestamp', value: 'Cumulative Water Discharge(Ltrs)' },
};

const GREEN = 'FFA9D18E';
const RED = 'FFC00000';
const THIN = { style: 'thin', color: { argb: 'FF000000' } };
const BORDER = { top: THIN, bottom: THIN, left: THIN, right: THIN };

const imeiOf = (v) => (v === null || v === undefined ? '' : String(v).trim());

function pickCol(table, wanted, label) {
  const hit = findHeader(table.headers, wanted);
  if (!hit) {
    throw new UserError(`${label}: column "${wanted}" not found. Available columns: ${table.headers.join(', ')}`);
  }
  return hit;
}

function withBlankRowsBetweenImeis(rows) {
  const out = [];
  let prev = null;
  for (const r of rows) {
    const cur = imeiOf(r.IMEI);
    if (prev !== null && cur !== prev) out.push(null);
    out.push(r);
    prev = cur;
  }
  return out;
}

function finishSheet(ws, { freeze = true } = {}) {
  ws.getRow(1).eachCell((c) => {
    c.font = { bold: true };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws.views = freeze ? [{ state: 'frozen', xSplit: 1, ySplit: 1 }] : [];
  const cols = ws.columnCount;
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= cols; c++) row.getCell(c).border = BORDER;
  }
}

async function matchTimestamps(ex1File, ex2File, opts = {}) {
  const defined = (o) => Object.fromEntries(Object.entries(o || {}).filter(([, v]) => v));
  const c1 = { ...DEFAULTS.ex1, ...defined(opts.columns?.ex1) };
  const c2 = { ...DEFAULTS.ex2, ...defined(opts.columns?.ex2) };

  const t1 = await readTable(ex1File.buffer, ex1File.name, { keys: [c1.imei] });
  const t2 = await readTable(ex2File.buffer, ex2File.name, { keys: [c2.imei] });

  const e1 = { imei: pickCol(t1, c1.imei, 'Excel 1'), ts: pickCol(t1, c1.timestamp, 'Excel 1'), val: pickCol(t1, c1.value, 'Excel 1') };
  const e2 = { imei: pickCol(t2, c2.imei, 'Excel 2'), ts: pickCol(t2, c2.timestamp, 'Excel 2'), val: pickCol(t2, c2.value, 'Excel 2') };

  // IMEI filter (empty list = every IMEI found in Excel 1)
  const wantedImeis = (opts.imeis || []).map(imeiOf).filter(Boolean);
  const imeis = wantedImeis.length ? [...new Set(wantedImeis)] : [...new Set(t1.rows.map((r) => imeiOf(r[e1.imei])).filter(Boolean))];
  const imeiSet = new Set(imeis);

  const from = limitKey(opts.fromDate, false);
  const to = limitKey(opts.toDate, true);
  if (opts.fromDate && !from) throw new UserError(`Invalid From date "${opts.fromDate}". Use YYYY-MM-DD.`);
  if (opts.toDate && !to) throw new UserError(`Invalid To date "${opts.toDate}". Use YYYY-MM-DD.`);
  const inRange = (k) => k && (!from || k >= from) && (!to || k <= to);

  const f1 = t1.rows.filter((r) => imeiSet.has(imeiOf(r[e1.imei])) && inRange(tsKey(r[e1.ts])));
  const f2 = t2.rows.filter((r) => imeiSet.has(imeiOf(r[e2.imei])) && inRange(tsKey(r[e2.ts])));

  const map = new Map(); // "imei|timestamp" -> Excel 2 rows
  for (const r of f2) {
    const k = tsKey(r[e2.ts]);
    if (!k) continue;
    const key = `${imeiOf(r[e2.imei])}|${k}`;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(r);
  }

  const matches = [];
  const missing = [];
  const extraHeaders = [];
  const addExtra = (h) => extraHeaders.includes(h) || extraHeaders.push(h);

  for (const r of f1) {
    const k = tsKey(r[e1.ts]);
    if (!k) continue;
    const hits = map.get(`${imeiOf(r[e1.imei])}|${k}`);
    if (!hits) {
      missing.push({ 'Excel1 Timestamp': show(r[e1.ts]), IMEI: imeiOf(r[e1.imei]), POPTOTWD1: r[e1.val], Status: 'Missing in Excel2' });
      continue;
    }
    for (const h of hits) {
      const row = {
        'Excel1 Timestamp': show(r[e1.ts]),
        'Excel2 Timestamp': show(h[e2.ts]),
        IMEI: imeiOf(r[e1.imei]),
        'Cumulative Water Discharge(Ltrs)': h[e2.val],
        POPTOTWD1: r[e1.val],
      };
      if (opts.includeAllColumns) {
        for (const [c, v] of Object.entries(r)) (row[`EX1_${c}`] = v), addExtra(`EX1_${c}`);
        for (const [c, v] of Object.entries(h)) (row[`EX2_${c}`] = v), addExtra(`EX2_${c}`);
      }
      matches.push(row);
    }
  }

  const wb = new ExcelJS.Workbook();

  const matchHeaders = ['Excel1 Timestamp', 'Excel2 Timestamp', 'IMEI', 'Cumulative Water Discharge(Ltrs)', 'POPTOTWD1', ...extraHeaders];
  const wsMatches = addTableSheet(wb, 'Matches', matchHeaders, withBlankRowsBetweenImeis(matches), {
    widths: [22, 24, 18, 28, 14, ...extraHeaders.map(() => 18)],
  });
  finishSheet(wsMatches);

  const wsMissing = addTableSheet(wb, 'Missing in Ex2', ['Excel1 Timestamp', 'IMEI', 'POPTOTWD1', 'Status'], withBlankRowsBetweenImeis(missing), {
    widths: [22, 18, 14, 20],
  });
  finishSheet(wsMissing);

  finishSheet(buildDateStatusSheet(wb, { f1, e1, imeis, map, from, to }));

  const stats = {
    'Excel 1 rows (after filter)': f1.length,
    'Excel 2 rows (after filter)': f2.length,
    'Exact matches': matches.length,
    'Missing in Excel 2': missing.length,
  };
  return {
    buffer: await toBuffer(wb),
    filename: 'timestamp_match_output.xlsx',
    summary: {
      stats,
      notes: [`IMEIs processed: ${imeis.length}`, `Date range: ${from ? from.slice(0, 10) : 'start'} to ${to ? to.slice(0, 10) : 'end'}`],
    },
  };
}

function buildDateStatusSheet(wb, { f1, e1, imeis, map, from, to }) {
  const daily = new Map();
  const dates = [];

  for (const r of f1) {
    const ts = tsKey(r[e1.ts]);
    if (!ts) continue;
    const imei = imeiOf(r[e1.imei]);
    const date = ts.slice(0, 10);
    dates.push(date);
    const key = `${imei}|${date}`;
    if (!daily.has(key)) daily.set(key, { matched: 0, missing: 0 });
    const s = daily.get(key);
    if (map.has(`${imei}|${ts}`)) s.matched++;
    else s.missing++;
  }

  const sorted = [...dates].sort();
  const startDate = from ? from.slice(0, 10) : sorted[0];
  const endDate = to ? to.slice(0, 10) : sorted[sorted.length - 1];

  const ws = wb.addWorksheet('My Color Sheet');
  ws.addRow(['Date', ...imeis]);
  ws.getColumn(1).width = 16;
  imeis.forEach((_, i) => (ws.getColumn(i + 2).width = 23));

  if (startDate && endDate) {
    const end = new Date(`${endDate}T00:00:00Z`);
    for (let d = new Date(`${startDate}T00:00:00Z`); d <= end; d.setUTCDate(d.getUTCDate() + 1)) {
      const dateKey = d.toISOString().slice(0, 10);
      const row = ws.addRow([new Date(d)]);
      const dateCell = row.getCell(1);
      dateCell.numFmt = 'dd/mm/yy';
      dateCell.font = { bold: true };
      dateCell.alignment = { horizontal: 'center', vertical: 'middle' };
      imeis.forEach((imei, j) => {
        const s = daily.get(`${imei}|${dateKey}`);
        if (!s) return;
        row.getCell(j + 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: s.missing === 0 ? GREEN : RED } };
      });
    }
  }
  return ws;
}

module.exports = { matchTimestamps, DEFAULTS };
