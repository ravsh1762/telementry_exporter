"use client";

import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  DEVICE_OPTIONS,
  FormValues,
  HP_OPTIONS,
  INITIAL_FORM_VALUES,
  STATE_OPTIONS,
  TELEMETRY_COLUMNS,
  TelemetryRecord,
  fetchTelemetryData,
  validateForm,
} from "./telemetry";

/* ------------------------------------------------------------------ */
/*  Template / column definition                                       */
/*  NOTE: the calculation/generation logic itself is untouched — this  */
/*  file only adds an Excel-based way of supplying the SAME inputs     */
/*  that the manual form already sends to /telemetry/generate.         */
/* ------------------------------------------------------------------ */

export interface BulkTemplateColumn {
  header: string;
  field: keyof FormValues;
  required: boolean;
  example: string;
  example2?: string;
  notes: string;
}

export const BULK_TEMPLATE_COLUMNS: BulkTemplateColumn[] = [
  {
    header: "IMEI",
    field: "imei",
    required: true,
    example: "864202050123456",
    example2: "864202050998877",
    notes: "Device IMEI number (numeric).",
  },
  {
    header: "Serial Number",
    field: "serialNumber",
    required: false,
    example: "SN-2024-001",
    example2: "SN-2024-002",
    notes: "Alphanumeric serial / ASN_11 code. Optional.",
  },
  {
    header: "HP Rating",
    field: "hp",
    required: true,
    example: "5",
    example2: "3",
    notes: `Allowed values: ${HP_OPTIONS.map((o) => o.value).join(", ")}`,
  },
  {
    header: "Device Type",
    field: "deviceType",
    required: true,
    example: "AC Surface",
    example2: "DC Submersible",
    notes: `Allowed values: ${DEVICE_OPTIONS.map((o) => o.value).join(", ")}`,
  },
  {
    header: "State",
    field: "state",
    required: false,
    example: "Rajasthan",
    example2: "Maharashtra",
    notes: `Allowed values: ${STATE_OPTIONS.map((o) => o.value).join(", ")}. POTP/COTP are set to a fixed value only when State is Maharashtra; blank for all other states.`,
  },
  {
    header: "From Date",
    field: "fromDate",
    required: true,
    example: "2026-07-01",
    example2: "2026-07-01",
    notes: "Format YYYY-MM-DD.",
  },
  {
    header: "To Date",
    field: "toDate",
    required: false,
    example: "",
    example2: "2026-07-05",
    notes:
      "Format YYYY-MM-DD. Leave blank to generate multiple records within From Date only.",
  },
  {
    header: "Record Count",
    field: "dataCount",
    required: false,
    example: "500",
    example2: "",
    notes:
      "Required only when To Date is blank or equal to From Date (max 100000). Ignored when a To Date range is given — one record per day is generated instead.",
  },
  {
    header: "Total Energy",
    field: "totalEnergy",
    required: false,
    example: "1200",
    example2: "800",
    notes: "Starting cumulative energy in kWh. Optional, defaults to 0.",
  },
  {
    header: "Total Water Discharge",
    field: "totalWaterDischarge",
    required: false,
    example: "50000",
    example2: "20000",
    notes:
      "Starting cumulative water discharge in litres. Optional, defaults to 0.",
  },
  {
    header: "Total Run Hours",
    field: "totalRunHours",
    required: false,
    example: "800",
    example2: "300",
    notes: "Starting cumulative run hours. Optional, defaults to 0.",
  },
];

const HEADER_ALIASES: Record<string, keyof FormValues> = {
  imei: "imei",
  serialnumber: "serialNumber",
  serialno: "serialNumber",
  serial: "serialNumber",
  asn11: "serialNumber",
  hp: "hp",
  hprating: "hp",
  hppower: "hp",
  devicetype: "deviceType",
  device: "deviceType",
  state: "state",
  statename: "state",
  region: "state",
  fromdate: "fromDate",
  startdate: "fromDate",
  todate: "toDate",
  enddate: "toDate",
  recordcount: "dataCount",
  datacount: "dataCount",
  count: "dataCount",
  totalenergy: "totalEnergy",
  totalenergykwh: "totalEnergy",
  totalwaterdischarge: "totalWaterDischarge",
  totalwaterdischargel: "totalWaterDischarge",
  totalwaterdischargeliters: "totalWaterDischarge",
  totalrunhours: "totalRunHours",
  runhours: "totalRunHours",
};

function normalizeHeader(header: unknown): string {
  return String(header ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function pad(n: number) {
  return String(n).padStart(2, "0");
}

function normalizeDateCell(value: unknown): string {
  if (value === null || value === undefined || value === "") return "";

  if (value instanceof Date && !isNaN(value.getTime())) {
    return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
  }

  const str = String(value).trim();
  if (!str) return "";

  // Already ISO-like (YYYY-MM-DD or YYYY/MM/DD)
  const isoMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    const [, y, m, d] = isoMatch;
    return `${y}-${pad(Number(m))}-${pad(Number(d))}`;
  }

  // DD-MM-YYYY or DD/MM/YYYY
  const dmyMatch = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (dmyMatch) {
    const [, d, m, y] = dmyMatch;
    return `${y}-${pad(Number(m))}-${pad(Number(d))}`;
  }

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
  }

  return str;
}

export interface BulkRow {
  rowId: string;
  rowNumber: number;
  values: FormValues;
  errors: Record<string, string>;
}

export interface BulkParseResult {
  rows: BulkRow[];
  error?: string;
}

export async function parseBulkFile(file: File): Promise<BulkParseResult> {
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });

    // Prefer a sheet literally called "Template" (matches the sample file);
    // otherwise fall back to the first sheet in the workbook.
    const sheetName =
      workbook.SheetNames.find((n) => n.trim().toLowerCase() === "template") ||
      workbook.SheetNames[0];

    if (!sheetName) {
      return { rows: [], error: "The uploaded file has no sheets." };
    }

    const sheet = workbook.Sheets[sheetName];
    const aoa: unknown[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: "",
      dateNF: "yyyy-mm-dd",
    });

    if (aoa.length === 0) {
      return { rows: [], error: "The sheet is empty." };
    }

    const headerRow = aoa[0];
    const fieldByColumn: (keyof FormValues | null)[] = headerRow.map((h) => {
      const normalized = normalizeHeader(h);
      return HEADER_ALIASES[normalized] ?? null;
    });

    if (!fieldByColumn.includes("imei") && !fieldByColumn.includes("fromDate")) {
      return {
        rows: [],
        error:
          "Could not recognize the column headers. Please use the downloaded sample template.",
      };
    }

    const rows: BulkRow[] = [];

    for (let i = 1; i < aoa.length; i++) {
      const rawRow = aoa[i];
      const isBlank =
        !rawRow ||
        rawRow.every((cell) => cell === "" || cell === null || cell === undefined);
      if (isBlank) continue;

      const values: FormValues = { ...INITIAL_FORM_VALUES };

      fieldByColumn.forEach((field, colIndex) => {
        if (!field) return;
        const raw = rawRow[colIndex];

        if (field === "fromDate" || field === "toDate") {
          values[field] = normalizeDateCell(raw);
        } else if (field === "hp") {
          values.hp = String(raw ?? "").trim();
        } else {
          values[field] = raw === null || raw === undefined ? "" : String(raw).trim();
        }
      });

      const errors = validateBulkRow(values);

      rows.push({
        rowId: `row-${i}-${Math.random().toString(36).slice(2, 8)}`,
        rowNumber: i + 1, // 1-based, accounting for header row
        values,
        errors,
      });
    }

    if (rows.length === 0) {
      return { rows: [], error: "No data rows found below the header row." };
    }

    return { rows };
  } catch (err) {
    return {
      rows: [],
      error:
        err instanceof Error
          ? `Failed to read Excel file: ${err.message}`
          : "Failed to read Excel file.",
    };
  }
}

export function validateBulkRow(values: FormValues): Record<string, string> {
  const errors: Record<string, string> = { ...validateForm(values) };

  if (!values.imei) {
    errors.imei = "IMEI is required";
  }

  if (!values.deviceType) {
    errors.deviceType = "Device type is required";
  } else if (!DEVICE_OPTIONS.some((o) => o.value === values.deviceType)) {
    errors.deviceType = `Must be one of: ${DEVICE_OPTIONS.map((o) => o.value).join(", ")}`;
  }

  if (values.hp && !HP_OPTIONS.some((o) => o.value === values.hp)) {
    errors.hp = `Must be one of: ${HP_OPTIONS.map((o) => o.value).join(", ")}`;
  }

  if (values.state && !STATE_OPTIONS.some((o) => o.value === values.state)) {
    errors.state = `Must be one of: ${STATE_OPTIONS.map((o) => o.value).join(", ")}`;
  }

  return errors;
}

/* ------------------------------------------------------------------ */
/*  Sample template download                                           */
/* ------------------------------------------------------------------ */

export function downloadSampleTemplate() {
  const headers = BULK_TEMPLATE_COLUMNS.map((c) => c.header);

  // One example row per state, so every valid state name is visible
  // directly in the sample sheet and easy to copy/select from.
  const hpCycle = HP_OPTIONS.map((o) => o.value);
  const deviceCycle = DEVICE_OPTIONS.map((o) => o.value);
  const baseImei = 864202050123456;

  const exampleRows = STATE_OPTIONS.map((stateOption, i) =>
    BULK_TEMPLATE_COLUMNS.map((c) => {
      switch (c.field) {
        case "imei":
          return String(baseImei + i);
        case "serialNumber":
          return `SN-2024-${String(i + 1).padStart(3, "0")}`;
        case "hp":
          return hpCycle[i % hpCycle.length];
        case "deviceType":
          return deviceCycle[i % deviceCycle.length];
        case "state":
          return stateOption.value;
        case "fromDate":
          return "2026-07-01";
        case "toDate":
          return "2026-07-05";
        case "dataCount":
          return "500";
        case "totalEnergy":
          return "1200";
        case "totalWaterDischarge":
          return "50000";
        case "totalRunHours":
          return "800";
        default:
          return c.example;
      }
    }),
  );

  const templateSheet = XLSX.utils.aoa_to_sheet([headers, ...exampleRows]);
  templateSheet["!cols"] = headers.map(() => ({ wch: 22 }));

  const refData: (string | number)[][] = [
    ["Field", "Required", "Notes / Allowed Values"],
    ...BULK_TEMPLATE_COLUMNS.map((c) => [
      c.header,
      c.required ? "Yes" : "No",
      c.notes,
    ]),
    ["", "", ""],
    [
      "Rule",
      "",
      "Provide either a Record Count (single date burst) OR a To Date (one record per day across a range). If To Date differs from From Date, Record Count is ignored.",
    ],
    [
      "Rule",
      "",
      "Replace the sample rows with your own devices, or add more rows below them. Do not rename the header row.",
    ],
    [
      "All States",
      "",
      STATE_OPTIONS.map((o) => o.value).join(", "),
    ],
  ];
  const refSheet = XLSX.utils.aoa_to_sheet(refData);
  refSheet["!cols"] = [{ wch: 24 }, { wch: 10 }, { wch: 85 }];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, templateSheet, "Template");
  XLSX.utils.book_append_sheet(workbook, refSheet, "Field Reference");

  XLSX.writeFile(workbook, "telemetry_bulk_upload_template.xlsx");
}

/* ------------------------------------------------------------------ */
/*  Generation (loops over rows, reusing the existing API call)        */
/* ------------------------------------------------------------------ */

export interface BulkResult {
  rowId: string;
  rowNumber: number;
  imei: string;
  serialNumber: string;
  deviceType: string;
  hp: string;
  status: "success" | "error";
  records: TelemetryRecord[];
  errorMessage?: string;
}

export async function runBulkGeneration(
  rows: BulkRow[],
  onProgress?: (completed: number, total: number, latest: BulkResult) => void,
): Promise<BulkResult[]> {
  const validRows = rows.filter((r) => Object.keys(r.errors).length === 0);
  const results: BulkResult[] = [];

  for (let i = 0; i < validRows.length; i++) {
    const row = validRows[i];
    try {
      const records = await fetchTelemetryData(row.values);
      const result: BulkResult = {
        rowId: row.rowId,
        rowNumber: row.rowNumber,
        imei: row.values.imei,
        serialNumber: row.values.serialNumber,
        deviceType: row.values.deviceType,
        hp: row.values.hp,
        status: "success",
        records,
      };
      results.push(result);
      onProgress?.(i + 1, validRows.length, result);
    } catch (err) {
      const result: BulkResult = {
        rowId: row.rowId,
        rowNumber: row.rowNumber,
        imei: row.values.imei,
        serialNumber: row.values.serialNumber,
        deviceType: row.values.deviceType,
        hp: row.values.hp,
        status: "error",
        records: [],
        errorMessage: err instanceof Error ? err.message : "Generation failed",
      };
      results.push(result);
      onProgress?.(i + 1, validRows.length, result);
    }
  }

  return results;
}

/* ------------------------------------------------------------------ */
/*  Export: every row's output goes into its OWN Excel file, bundled   */
/*  together in a single ZIP download.                                 */
/* ------------------------------------------------------------------ */

function buildWorksheetBuffer(
  records: TelemetryRecord[],
  columnMap: Record<string, string>,
) {
  const fields = Object.keys(columnMap);
  const header = fields;
  const rows = records.map((row) =>
    fields.map((field) => (row as Record<string, unknown>)[field] ?? ""),
  );
  const sheetData = [header, ...rows];

  const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
  worksheet["!cols"] = fields.map(() => ({ wch: 18 }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Telemetry");

  return XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}

function sanitizeSheetName(name: string, usedNames: Set<string>): string {
  let base =
    String(name || "Device")
      .replace(/[\\/?*\[\]:]/g, "_")
      .trim()
      .substring(0, 28) || "Device";

  let candidate = base;
  let i = 1;
  while (usedNames.has(candidate.toLowerCase())) {
    candidate = `${base}_${i}`.substring(0, 31);
    i++;
  }
  usedNames.add(candidate.toLowerCase());
  return candidate;
}

/**
 * One Excel workbook with:
 *  - a "Master" sheet containing every generated record from every device
 *    (tagged with IMEI / Serial / Device Type / HP so rows are traceable)
 *  - one additional sheet per device, containing only that device's records
 */
export function exportBulkResultsAsWorkbook(results: BulkResult[]) {
  const columnMap = Object.fromEntries(
    TELEMETRY_COLUMNS.map((col) => [col.key, col.label]),
  );
  const fields = Object.keys(columnMap);

  const successResults = results.filter(
    (r) => r.status === "success" && r.records.length > 0,
  );
  if (successResults.length === 0) return;

  const workbook = XLSX.utils.book_new();

  // ---- Master sheet: every record, from every device, in one place ----
  // Uses the raw field keys (VD, TIMESTAMP, MAXINDEX, ...) as headers,
  // matching the API's field names exactly.
  const masterHeader = fields;
  const masterRows: (string | number)[][] = [];
  successResults.forEach((result, resultIndex) => {
    result.records.forEach((row) => {
      masterRows.push(
        fields.map((f) => (row as Record<string, unknown>)[f] as string | number ?? ""),
      );
    });

    // Blank spacer row between one device's block and the next
    if (resultIndex < successResults.length - 1) {
      masterRows.push(masterHeader.map(() => ""));
    }
  });
  const masterSheet = XLSX.utils.aoa_to_sheet([masterHeader, ...masterRows]);
  masterSheet["!cols"] = masterHeader.map(() => ({ wch: 18 }));
  XLSX.utils.book_append_sheet(workbook, masterSheet, "Master");

  // ---- One sheet per device ----
  const usedNames = new Set<string>(["master"]);
  successResults.forEach((result) => {
    const sheetName = sanitizeSheetName(
      result.imei || `Row${result.rowNumber}`,
      usedNames,
    );
    const header = fields.map((f) => columnMap[f]);
    const rows = result.records.map((row) =>
      fields.map((f) => (row as Record<string, unknown>)[f] ?? ""),
    );
    const sheet = XLSX.utils.aoa_to_sheet([header, ...rows]);
    sheet["!cols"] = header.map(() => ({ wch: 18 }));
    XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  });

  const timestamp = new Date()
    .toISOString()
    .replace(/[:.]/g, "-")
    .substring(0, 16);

  XLSX.writeFile(workbook, `telemetry_bulk_combined_${timestamp}.xlsx`);
}

export async function exportBulkResultsAsZip(results: BulkResult[]) {
  const columnMap = Object.fromEntries(
    TELEMETRY_COLUMNS.map((col) => [col.key, col.label]),
  );

  const successResults = results.filter(
    (r) => r.status === "success" && r.records.length > 0,
  );
  if (successResults.length === 0) return;

  const zip = new JSZip();
  const timestamp = new Date()
    .toISOString()
    .replace(/[:.]/g, "-")
    .substring(0, 16);

  successResults.forEach((result) => {
    const buffer = buildWorksheetBuffer(result.records, columnMap);
    const safeImei = (result.imei || `row${result.rowNumber}`).replace(
      /\s+/g,
      "_",
    );
    const safeDevice = (result.deviceType || "device").replace(/\s+/g, "_");
    const filename = `telemetry_${safeImei}_${safeDevice}_${timestamp}.xlsx`;
    zip.file(filename, buffer);
  });

  const zipBlob = await zip.generateAsync({ type: "blob" });
  const url = URL.createObjectURL(zipBlob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `telemetry_bulk_export_${timestamp}.zip`;
  link.click();
  URL.revokeObjectURL(url);
}

export function exportSingleResultAsExcel(result: BulkResult) {
  const columnMap = Object.fromEntries(
    TELEMETRY_COLUMNS.map((col) => [col.key, col.label]),
  );
  const buffer = buildWorksheetBuffer(result.records, columnMap);
  const blob = new Blob([buffer], {
    type: "application/octet-stream",
  });
  const timestamp = new Date()
    .toISOString()
    .replace(/[:.]/g, "-")
    .substring(0, 16);
  const safeImei = (result.imei || `row${result.rowNumber}`).replace(
    /\s+/g,
    "_",
  );
  const safeDevice = (result.deviceType || "device").replace(/\s+/g, "_");

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `telemetry_${safeImei}_${safeDevice}_${timestamp}.xlsx`;
  link.click();
  URL.revokeObjectURL(url);
}
