'use strict';
/** Date / timestamp parsing shared by the tools. */

const pad = (n) => String(n).padStart(2, '0');

const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

/**
 * Understands:
 *   2026-06-11 16:25:31   (also with "/" or "T")
 *   27 Sep 2026 18:08:52
 *   11/06/2026 , 06-11-2026   (if the first number is > 12 it is the day, otherwise the month)
 *   Excel date cells (Date objects / serial numbers)
 * @returns {{y,m,d,H,M,S,hasTime}|null}
 */
function parts(v) {
  if (v === null || v === undefined || v === '') return null;

  const fromDate = (d, local) =>
    Number.isNaN(d.getTime())
      ? null
      : local
        ? { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate(), H: d.getHours(), M: d.getMinutes(), S: d.getSeconds(), hasTime: true }
        : { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), H: d.getUTCHours(), M: d.getUTCMinutes(), S: d.getUTCSeconds(), hasTime: true };

  if (v instanceof Date) return fromDate(v, false); // ExcelJS gives UTC-based dates
  if (typeof v === 'number') return fromDate(new Date(Math.round(((v - 25569) * 86400000) / 1000) * 1000), false);

  const s = String(v).trim();
  if (!s) return null;
  const time = (a, b, c) => ({ H: Number(a ?? 0), M: Number(b ?? 0), S: Number(c ?? 0), hasTime: a !== undefined });

  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return { y: +m[1], m: +m[2], d: +m[3], ...time(m[4], m[5], m[6]) };

  m = s.match(/^(\d{1,2})\s+([A-Za-z]{3})[a-z]*\s+(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    const mon = MONTHS[m[2].toLowerCase()];
    if (!mon) return null;
    return { y: +m[3], m: mon, d: +m[1], ...time(m[4], m[5], m[6]) };
  }

  m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const a = +m[1];
    const b = +m[2];
    const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    const [mo, d] = a > 12 && b <= 12 ? [b, a] : [a, b];
    return { y, m: mo, d, ...time(m[4], m[5], m[6]) };
  }

  if (/^\d+$/.test(s)) return null; // e.g. "260521" - not a real date
  return fromDate(new Date(s), true);
}

/** "YYYY-MM-DD HH:MM:SS" (null when the value has no time part) */
function tsKey(v) {
  const p = parts(v);
  if (!p || !p.hasTime) return null;
  return `${p.y}-${pad(p.m)}-${pad(p.d)} ${pad(p.H)}:${pad(p.M)}:${pad(p.S)}`;
}

/** "YYYY-MM-DD" */
function dateKey(v) {
  const p = parts(v);
  return p ? `${p.y}-${pad(p.m)}-${pad(p.d)}` : null;
}

/** Turns a from/to limit into a comparable key. Empty -> null. */
function limitKey(v, isEnd) {
  const s = String(v || '').trim();
  if (!s) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s + (isEnd ? ' 23:59:59' : ' 00:00:00');
  return tsKey(s);
}

/** Show real Excel dates as readable text, leave everything else untouched. */
const show = (v) => (typeof v === 'number' || v instanceof Date ? tsKey(v) : v);

module.exports = { pad, parts, tsKey, dateKey, limitKey, show };
