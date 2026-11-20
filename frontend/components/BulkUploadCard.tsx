"use client";

import React, { useCallback, useRef, useState } from "react";
import { Button } from "./Formelements";
import {
  BulkResult,
  BulkRow,
  downloadSampleTemplate,
  exportBulkResultsAsWorkbook,
  exportBulkResultsAsZip,
  exportSingleResultAsExcel,
  parseBulkFile,
  runBulkGeneration,
} from "@/lib/bulkTelemetry";

export function BulkUploadCard() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [rows, setRows] = useState<BulkRow[]>([]);
  const [parseError, setParseError] = useState<string | null>(null);
  const [generating, setGenerating] = useState(false);
  const [progress, setProgress] = useState<{ done: number; total: number }>({
    done: 0,
    total: 0,
  });
  const [results, setResults] = useState<BulkResult[]>([]);

  const validRows = rows.filter((r) => Object.keys(r.errors).length === 0);
  const invalidRows = rows.filter((r) => Object.keys(r.errors).length > 0);

  const handleFileChange = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      setFileName(file.name);
      setResults([]);
      setParseError(null);
      setRows([]);

      const { rows: parsedRows, error } = await parseBulkFile(file);
      if (error) {
        setParseError(error);
        return;
      }
      setRows(parsedRows);
    },
    [],
  );

  const handleGenerateAll = useCallback(async () => {
    if (validRows.length === 0) return;

    setGenerating(true);
    setResults([]);
    setProgress({ done: 0, total: validRows.length });

    const finalResults = await runBulkGeneration(validRows, (completed, total) => {
      setProgress({ done: completed, total });
    });

    setResults(finalResults);
    setGenerating(false);
  }, [validRows]);

  const handleClear = useCallback(() => {
    setFileName(null);
    setRows([]);
    setParseError(null);
    setResults([]);
    setProgress({ done: 0, total: 0 });
    if (fileInputRef.current) fileInputRef.current.value = "";
  }, []);

  const successCount = results.filter((r) => r.status === "success").length;
  const errorCount = results.filter((r) => r.status === "error").length;
  const totalGeneratedRecords = results.reduce(
    (sum, r) => sum + r.records.length,
    0,
  );

  return (
    <div className="bg-white border border-slate-200 shadow-sm rounded-xl overflow-hidden">
      <div className="px-6 py-5 border-b border-slate-200 flex items-center justify-between bg-white flex-wrap gap-3">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-blue-100 border border-blue-200 flex items-center justify-center" />
          <div>
            <h2 className="text-sm font-semibold text-gray-700">
              Bulk Excel Upload
            </h2>
            <p className="text-xs text-gray-500 mt-0.5">
              Generate telemetry for many devices at once from a spreadsheet
            </p>
          </div>
        </div>

        <Button variant="secondary" size="sm" onClick={downloadSampleTemplate}>
          Download Sample Template
        </Button>
      </div>

      <div className="p-6 space-y-5 bg-white">
        <div>
          <p className="text-xs font-semibold tracking-widest uppercase text-gray-700 mb-3">
            1. Upload Filled Excel Sheet
          </p>

          <div className="flex items-center gap-3 flex-wrap">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileChange}
              className="hidden"
              id="bulk-file-input"
            />
            <label
              htmlFor="bulk-file-input"
              className="cursor-pointer inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold rounded-lg bg-slate-800 text-slate-200 border border-slate-200 hover:bg-slate-700 hover:border-slate-600 transition-all"
            >
              Choose Excel File
            </label>

            {fileName && (
              <span className="text-xs font-mono text-slate-500 bg-slate-50 border border-slate-200 rounded-md px-2.5 py-1.5">
                {fileName}
              </span>
            )}

            {(rows.length > 0 || parseError) && (
              <Button variant="ghost" size="sm" onClick={handleClear}>
                Clear
              </Button>
            )}
          </div>

          <p className="text-xs text-slate-400 mt-2">
            Required columns: IMEI, HP Rating, Device Type, From Date — plus
            either Record Count or a To Date. Use the sample template above
            for the exact column names.
          </p>
        </div>

        {parseError && (
          <div className="bg-red-50 border border-red-200 rounded-xl px-4 py-3 text-sm text-red-700">
            {parseError}
          </div>
        )}

        {rows.length > 0 && (
          <div>
            <div className="flex items-center gap-2 mb-3 flex-wrap">
              <span className="text-xs font-semibold tracking-widest uppercase text-gray-700">
                2. Preview
              </span>
              <span className="font-mono text-xs px-2 py-0.5 bg-slate-100 text-slate-600 border border-slate-200 rounded-md">
                {rows.length} rows
              </span>
              <span className="font-mono text-xs px-2 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-md">
                {validRows.length} valid
              </span>
              {invalidRows.length > 0 && (
                <span className="font-mono text-xs px-2 py-0.5 bg-red-50 text-red-700 border border-red-200 rounded-md">
                  {invalidRows.length} invalid
                </span>
              )}
            </div>

            <div className="border border-slate-200 rounded-lg overflow-auto max-h-64">
              <table className="w-full text-xs border-collapse">
                <thead className="sticky top-0 bg-slate-50 z-10">
                  <tr className="border-b border-slate-200">
                    <th className="text-left px-3 py-2 font-semibold text-slate-500">
                      Row
                    </th>
                    <th className="text-left px-3 py-2 font-semibold text-slate-500">
                      IMEI
                    </th>
                    <th className="text-left px-3 py-2 font-semibold text-slate-500">
                      Device
                    </th>
                    <th className="text-left px-3 py-2 font-semibold text-slate-500">
                      HP
                    </th>
                    <th className="text-left px-3 py-2 font-semibold text-slate-500">
                      State
                    </th>
                    <th className="text-left px-3 py-2 font-semibold text-slate-500">
                      From → To
                    </th>
                    <th className="text-left px-3 py-2 font-semibold text-slate-500">
                      Status
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row) => {
                    const hasErrors = Object.keys(row.errors).length > 0;
                    return (
                      <tr
                        key={row.rowId}
                        className="border-b border-slate-100 last:border-0"
                      >
                        <td className="px-3 py-2 text-slate-400 font-mono">
                          {row.rowNumber}
                        </td>
                        <td className="px-3 py-2 font-mono text-slate-700">
                          {row.values.imei || "--"}
                        </td>
                        <td className="px-3 py-2 text-slate-600">
                          {row.values.deviceType || "--"}
                        </td>
                        <td className="px-3 py-2 text-slate-600">
                          {row.values.hp || "--"}
                        </td>
                        <td className="px-3 py-2 text-slate-600">
                          {row.values.state || "--"}
                        </td>
                        <td className="px-3 py-2 text-slate-600 font-mono">
                          {row.values.fromDate || "--"}
                          {row.values.toDate ? ` → ${row.values.toDate}` : ""}
                        </td>
                        <td className="px-3 py-2">
                          {hasErrors ? (
                            <span
                              className="text-red-600 font-medium cursor-help"
                              title={Object.values(row.errors).join("; ")}
                            >
                              Invalid
                            </span>
                          ) : (
                            <span className="text-emerald-600 font-medium">
                              Ready
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      {rows.length > 0 && (
        <div className="px-6 py-4 border-t border-slate-200 bg-gray-50 flex items-center justify-between flex-wrap gap-3">
          <div className="text-xs text-slate-500">
            {generating
              ? `Generating ${progress.done}/${progress.total}...`
              : results.length > 0
                ? `${successCount} succeeded · ${errorCount} failed · ${totalGeneratedRecords.toLocaleString()} records total`
                : `${validRows.length} row(s) ready to generate`}
          </div>

          <div className="flex items-center gap-2 flex-wrap justify-end">
            {results.length > 0 && successCount > 0 && (
              <>
                <Button
                  variant="secondary"
                  size="sm"
                  onClick={() => exportBulkResultsAsWorkbook(results)}
                >
                  Download Combined Workbook (Master + per-device sheets)
                </Button>
                <Button
                  variant="secondary"
                  size="sm"
                  onClick={() => exportBulkResultsAsZip(results)}
                >
                  Download Separate Files (.zip)
                </Button>
              </>
            )}
            <Button
              variant="primary"
              onClick={handleGenerateAll}
              disabled={validRows.length === 0}
              loading={generating}
            >
              {generating ? "Generating..." : "Generate All Valid Rows"}
            </Button>
          </div>
        </div>
      )}

      {results.length > 0 && (
        <div className="px-6 py-4 border-t border-slate-200 bg-white">
          <p className="text-xs font-semibold tracking-widest uppercase text-gray-700 mb-3">
            3. Per-Device Results
          </p>
          <div className="border border-slate-200 rounded-lg overflow-auto max-h-64">
            <table className="w-full text-xs border-collapse">
              <thead className="sticky top-0 bg-slate-50 z-10">
                <tr className="border-b border-slate-200">
                  <th className="text-left px-3 py-2 font-semibold text-slate-500">
                    Row
                  </th>
                  <th className="text-left px-3 py-2 font-semibold text-slate-500">
                    IMEI
                  </th>
                  <th className="text-left px-3 py-2 font-semibold text-slate-500">
                    Device
                  </th>
                  <th className="text-left px-3 py-2 font-semibold text-slate-500">
                    Records
                  </th>
                  <th className="text-left px-3 py-2 font-semibold text-slate-500">
                    Status
                  </th>
                  <th className="text-left px-3 py-2 font-semibold text-slate-500">
                    Download
                  </th>
                </tr>
              </thead>
              <tbody>
                {results.map((result) => (
                  <tr
                    key={result.rowId}
                    className="border-b border-slate-100 last:border-0"
                  >
                    <td className="px-3 py-2 text-slate-400 font-mono">
                      {result.rowNumber}
                    </td>
                    <td className="px-3 py-2 font-mono text-slate-700">
                      {result.imei}
                    </td>
                    <td className="px-3 py-2 text-slate-600">
                      {result.deviceType}
                    </td>
                    <td className="px-3 py-2 text-slate-600 font-mono">
                      {result.records.length}
                    </td>
                    <td className="px-3 py-2">
                      {result.status === "success" ? (
                        <span className="text-emerald-600 font-medium">
                          Success
                        </span>
                      ) : (
                        <span
                          className="text-red-600 font-medium cursor-help"
                          title={result.errorMessage}
                        >
                          Failed
                        </span>
                      )}
                    </td>
                    <td className="px-3 py-2">
                      {result.status === "success" && (
                        <button
                          onClick={() => exportSingleResultAsExcel(result)}
                          className="text-blue-600 hover:text-blue-700 font-medium underline underline-offset-2"
                        >
                          .xlsx
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
