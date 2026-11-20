'use strict';
const { UserError } = require('../services/tools/io');
const { mergeReports } = require('../services/tools/reportMerger');
const { splitByImei } = require('../services/tools/imeiSplitter');
const { matchTimestamps } = require('../services/tools/timestampMatcher');
const { filterByDate } = require('../services/tools/dateFilter');
const { shiftDays } = require('../services/tools/dayShift');
const { extractLastImeiRows } = require('../services/tools/lastImei');

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// multer decodes file names as latin1; restore UTF-8
const fileName = (f) => Buffer.from(f.originalname, 'latin1').toString('utf8');
const asFile = (f) => ({ name: fileName(f), buffer: f.buffer });

function requireFile(f, label) {
  if (!f) throw new UserError(`Please choose ${label}.`);
  return asFile(f);
}

function send(res, { buffer, filename, contentType = XLSX_TYPE, summary }) {
  res.set({
    'Content-Type': contentType,
    'Content-Disposition': `attachment; filename="${filename}"`,
    'X-Tool-Summary': encodeURIComponent(JSON.stringify(summary || {})),
  });
  res.send(buffer);
}

/** wraps a handler: UserError -> 400, anything else -> 500 */
const handle = (fn) => async (req, res) => {
  try {
    send(res, await fn(req));
  } catch (err) {
    const status = err instanceof UserError ? 400 : 500;
    if (status === 500) console.error(err);
    res.status(status).json({ success: false, error: err.message });
  }
};

const int = (v, d) => (v === undefined || v === '' ? d : Number.parseInt(v, 10));

const reportMerger = handle((req) => {
  const files = (req.files || []).map(asFile);
  return mergeReports(files);
});

const imeiSplitter = handle((req) => {
  const f = requireFile(req.file, 'an Excel file');
  return splitByImei(f.buffer, f.name, {
    imeiCount: int(req.body.imeiCount, 0),
    headerRow: int(req.body.headerRow, 1) || 1,
    imeiColumn: req.body.imeiColumn || 'IMEI',
  });
});

const timestampMatcher = handle((req) => {
  const ex1 = requireFile(req.files?.ex1?.[0], 'Excel 1');
  const ex2 = requireFile(req.files?.ex2?.[0], 'Excel 2');
  const b = req.body;
  const col = (k) => (b[k] || '').trim() || undefined;
  return matchTimestamps(ex1, ex2, {
    imeis: String(b.imeis || '').split(/[\s,;]+/),
    fromDate: b.fromDate,
    toDate: b.toDate,
    includeAllColumns: b.includeAllColumns === 'true',
    columns: {
      ex1: { imei: col('ex1Imei'), timestamp: col('ex1Timestamp'), value: col('ex1Value') },
      ex2: { imei: col('ex2Imei'), timestamp: col('ex2Timestamp'), value: col('ex2Value') },
    },
  });
});

const dateFilter = handle((req) => {
  const f = requireFile(req.file, 'an Excel file');
  return filterByDate(f.buffer, f.name, { fromDate: req.body.fromDate, toDate: req.body.toDate });
});

const dayShift = handle((req) => {
  const f = requireFile(req.file, 'an Excel file');
  return shiftDays(f.buffer, f.name, { days: int(req.body.days, 1), keepTime: req.body.keepTime === 'true' });
});

const lastImei = handle((req) => {
  const f = requireFile(req.file, 'an Excel file');
  return extractLastImeiRows(f.buffer, f.name);
});

module.exports = { reportMerger, imeiSplitter, timestampMatcher, dateFilter, dayShift, lastImei };
