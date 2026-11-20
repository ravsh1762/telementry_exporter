# Telemetry Exporter

One app, seven tools. The original **Telemetry Generator** plus six Excel tools
that used to be separate projects.

| Page | What it does | Replaces |
|---|---|---|
| Telemetry Generator (`/`) | Generates pump telemetry (manual form or bulk Excel upload) | *(original)* |
| Report Merger | Merges many "Report Table" exports into one sheet (title row removed, blank row between files) + an IMEI1, IMEI2… list on Sheet2 | `pump-report-merger` |
| Last IMEI | Keeps only the last row of every IMEI | `Excel_Last-IMEI-Extractor` |
| Date Filter | Keeps rows inside a date range, grouped by IMEI with blank rows between | `excel-date-filter-project` |
| Day Shift | Adds N days to every TIMESTAMP | `1Day++ Date Excel project` |
| IMEI Splitter | Splits one sheet into EXCEL_1.xlsx, EXCEL_2.xlsx … (ZIP) | `imei-splitter` |
| Timestamp Matcher | Matches Excel 1 vs Excel 2 by IMEI + exact timestamp (Matches, Missing, colour sheet) | `timestamp-matcher-js` |

## Run it

Needs Node.js 18+ (20+ recommended).

```bash
npm install     # installs root + backend + frontend in one go
npm run dev     # starts backend (port 4000) and frontend (port 3000)
```

Open **http://localhost:3000** and use the menu at the top.

For a production-style run: `npm run build` then `npm start`.

## Project layout

```
backend/                      Express API (port 4000)
  index.js
  routes/routes.js            POST /telemetry/generate        (generator)
  routes/toolsRoutes.js       POST /tools/<tool>              (Excel tools)
  controller/toolsController.js
  services/tools/             one file per tool + shared io.js / timeUtils.js
frontend/                     Next.js app (port 3000)
  app/page.tsx                Telemetry Generator
  app/tools/<tool>/page.tsx   one page per Excel tool
  components/SiteHeader.tsx   shared menu
  components/tools/           upload box, result panel, form helpers
  lib/api.ts                  backend URL + download helper
```

## Notes

* Excel tools accept **.xlsx** only (not old .xls).
* All tools read/write with a single library (ExcelJS); the old `xlsx` / `xlsx-js-style` packages are no longer needed.
* Junk title rows above the header (like "Report Table") are skipped automatically: the header row is the first row that contains `IMEI` (or `TIMESTAMP` for the date tools).
* Timestamp Matcher: leave the IMEI list empty to process every IMEI found in Excel 1.
* Day Shift writes date only (YYYY-MM-DD) like the original script; tick "Keep the time of day" to keep HH:MM:SS.
* Backend URL: the frontend calls `http://localhost:4000`. To change it, set `NEXT_PUBLIC_API_URL` before running the frontend (the generator's own URL is in `frontend/lib/telemetry.ts`).
