"use client";

import { useState } from "react";
import { FileDrop } from "@/components/tools/FileDrop";
import { Card, ResultPanel, RunButton, ToolPage } from "@/components/tools/ui";
import { useToolRunner } from "@/lib/useToolRunner";

export default function LastImeiPage() {
  const [files, setFiles] = useState<File[]>([]);
  const { loading, error, result, run, fail } = useToolRunner();

  const submit = () => {
    if (!files.length) return fail("Please choose an .xlsx file.");
    const form = new FormData();
    form.append("file", files[0]);
    run("/tools/last-imei", form);
  };

  return (
    <ToolPage
      title="Last IMEI Extractor"
      description="Keeps only the last row of every IMEI. Output sheet: “Last IMEI Rows”."
    >
      <Card>
        <FileDrop label="Excel file" files={files} onChange={setFiles} />
        <RunButton loading={loading} disabled={!files.length} onClick={submit}>
          Extract last rows
        </RunButton>
      </Card>
      <ResultPanel result={result} error={error} />
    </ToolPage>
  );
}
