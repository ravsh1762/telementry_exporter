"use client";

import { useState } from "react";
import { FileDrop } from "@/components/tools/FileDrop";
import {
  Card,
  Field,
  inputClass,
  ResultPanel,
  RunButton,
  ToolPage,
} from "@/components/tools/ui";
import { useToolRunner } from "@/lib/useToolRunner";

export default function DateFilterPage() {
  const [files, setFiles] = useState<File[]>([]);
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const { loading, error, result, run, fail } = useToolRunner();

  const submit = () => {
    if (!files.length) return fail("Please choose an .xlsx file.");
    if (!fromDate) return fail("Please pick a From date.");
    if (toDate && toDate < fromDate) return fail("To date must be on or after From date.");
    const form = new FormData();
    form.append("file", files[0]);
    form.append("fromDate", fromDate);
    form.append("toDate", toDate || fromDate);
    run("/tools/date-filter", form);
  };

  return (
    <ToolPage
      title="Date Filter"
      description="Keeps only the rows whose TIMESTAMP is inside the date range (inclusive), grouped by IMEI with a blank row between IMEIs."
    >
      <Card>
        <FileDrop label="Excel file" files={files} onChange={setFiles} />
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="From date">
            <input
              type="date"
              value={fromDate}
              onChange={(e) => setFromDate(e.target.value)}
              className={inputClass}
            />
          </Field>
          <Field label="To date" hint="Leave empty for a single day">
            <input
              type="date"
              value={toDate}
              onChange={(e) => setToDate(e.target.value)}
              className={inputClass}
            />
          </Field>
        </div>
        <RunButton loading={loading} disabled={!files.length} onClick={submit}>
          Filter by date
        </RunButton>
      </Card>
      <ResultPanel result={result} error={error} />
    </ToolPage>
  );
}
