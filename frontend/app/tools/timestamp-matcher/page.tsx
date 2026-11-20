"use client";

import { useState } from "react";
import { FileDrop } from "@/components/tools/FileDrop";
import {
  Card,
  Checkbox,
  Field,
  inputClass,
  ResultPanel,
  RunButton,
  ToolPage,
} from "@/components/tools/ui";
import { useToolRunner } from "@/lib/useToolRunner";

const COLUMN_FIELDS = [
  { key: "ex1Imei", label: "Excel 1 · IMEI", placeholder: "IMEI" },
  { key: "ex1Timestamp", label: "Excel 1 · Timestamp", placeholder: "TIMESTAMP" },
  { key: "ex1Value", label: "Excel 1 · Value", placeholder: "POPTOTWD1" },
  { key: "ex2Imei", label: "Excel 2 · IMEI", placeholder: "IMEI" },
  { key: "ex2Timestamp", label: "Excel 2 · Timestamp", placeholder: "timestamp" },
  {
    key: "ex2Value",
    label: "Excel 2 · Value",
    placeholder: "Cumulative Water Discharge(Ltrs)",
  },
] as const;

export default function TimestampMatcherPage() {
  const [ex1, setEx1] = useState<File[]>([]);
  const [ex2, setEx2] = useState<File[]>([]);
  const [imeis, setImeis] = useState("");
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [includeAll, setIncludeAll] = useState(false);
  const [columns, setColumns] = useState<Record<string, string>>({});
  const { loading, error, result, run, fail } = useToolRunner();

  const submit = () => {
    if (!ex1.length || !ex2.length) return fail("Please choose both Excel 1 and Excel 2.");
    if (fromDate && toDate && toDate < fromDate)
      return fail("To date must be on or after From date.");
    const form = new FormData();
    form.append("ex1", ex1[0]);
    form.append("ex2", ex2[0]);
    form.append("imeis", imeis);
    form.append("fromDate", fromDate);
    form.append("toDate", toDate);
    form.append("includeAllColumns", String(includeAll));
    COLUMN_FIELDS.forEach(({ key }) => form.append(key, columns[key] ?? ""));
    run("/tools/timestamp-matcher", form);
  };

  return (
    <ToolPage
      title="Timestamp Matcher"
      description="Matches Excel 1 rows against Excel 2 rows by IMEI + exact timestamp. Output sheets: Matches, Missing in Ex2, and a date-wise colour sheet (green = all timestamps matched, red = something missing)."
    >
      <Card>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
          <FileDrop
            label="Excel 1 (TIMESTAMP like 2026-06-11 16:25:31)"
            files={ex1}
            onChange={setEx1}
          />
          <FileDrop
            label="Excel 2 (timestamp like 27 Sep 2026 18:08:52)"
            files={ex2}
            onChange={setEx2}
          />
        </div>

        <Field
          label="IMEI list"
          hint="One per line, or separated by spaces / commas. Leave empty to use every IMEI found in Excel 1."
        >
          <textarea
            rows={5}
            value={imeis}
            onChange={(e) => setImeis(e.target.value)}
            placeholder={"864524072136165\n864524072116001"}
            className={`${inputClass} font-mono`}
          />
        </Field>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="From date" hint="Optional">
            <input
              type="date"
              value={fromDate}
              onChange={(e) => setFromDate(e.target.value)}
              className={inputClass}
            />
          </Field>
          <Field label="To date" hint="Optional">
            <input
              type="date"
              value={toDate}
              onChange={(e) => setToDate(e.target.value)}
              className={inputClass}
            />
          </Field>
        </div>

        <Checkbox
          checked={includeAll}
          onChange={setIncludeAll}
          label="Also add every column of Excel 1 and Excel 2 to the Matches sheet"
        />

        <details className="group">
          <summary className="text-sm font-semibold text-slate-500 cursor-pointer select-none">
            Column names (only change if your files use different headers)
          </summary>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-4">
            {COLUMN_FIELDS.map(({ key, label, placeholder }) => (
              <Field key={key} label={label}>
                <input
                  value={columns[key] ?? ""}
                  placeholder={placeholder}
                  onChange={(e) =>
                    setColumns((prev) => ({ ...prev, [key]: e.target.value }))
                  }
                  className={inputClass}
                />
              </Field>
            ))}
          </div>
        </details>

        <RunButton
          loading={loading}
          disabled={!ex1.length || !ex2.length}
          onClick={submit}
        >
          Match timestamps
        </RunButton>
      </Card>
      <ResultPanel result={result} error={error} />
    </ToolPage>
  );
}
