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

export default function ImeiSplitterPage() {
  const [files, setFiles] = useState<File[]>([]);
  const [imeiColumn, setImeiColumn] = useState("IMEI");
  const [imeiCount, setImeiCount] = useState("2");
  const [headerRow, setHeaderRow] = useState("1");
  const { loading, error, result, run, fail } = useToolRunner();

  const submit = () => {
    if (!files.length) return fail("Please choose an .xlsx file.");
    if (!(Number(imeiCount) >= 1)) return fail("IMEI count must be 1 or more.");
    const form = new FormData();
    form.append("file", files[0]);
    form.append("imeiColumn", imeiColumn);
    form.append("imeiCount", imeiCount);
    form.append("headerRow", headerRow);
    run("/tools/imei-splitter", form);
  };

  return (
    <ToolPage
      title="IMEI Splitter"
      description="Splits one sheet into several Excel files (EXCEL_1.xlsx, EXCEL_2.xlsx …) with a fixed number of IMEIs in each, and downloads them as a ZIP. Original rows and formatting are kept; a blank IMEI cell belongs to the IMEI above it."
    >
      <Card>
        <FileDrop label="Excel file" files={files} onChange={setFiles} />
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          <Field label="IMEI column name">
            <input
              value={imeiColumn}
              onChange={(e) => setImeiColumn(e.target.value)}
              className={inputClass}
            />
          </Field>
          <Field label="IMEIs per file">
            <input
              type="number"
              min={1}
              value={imeiCount}
              onChange={(e) => setImeiCount(e.target.value)}
              className={inputClass}
            />
          </Field>
          <Field label="Header row number">
            <input
              type="number"
              min={1}
              value={headerRow}
              onChange={(e) => setHeaderRow(e.target.value)}
              className={inputClass}
            />
          </Field>
        </div>
        <RunButton loading={loading} disabled={!files.length} onClick={submit}>
          Split &amp; download ZIP
        </RunButton>
      </Card>
      <ResultPanel result={result} error={error} />
    </ToolPage>
  );
}
