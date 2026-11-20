"use client";

import { useState } from "react";
import { FileDrop } from "@/components/tools/FileDrop";
import { Card, ResultPanel, RunButton, ToolPage } from "@/components/tools/ui";
import { useToolRunner } from "@/lib/useToolRunner";

export default function ReportMergerPage() {
  const [files, setFiles] = useState<File[]>([]);
  const { loading, error, result, run, fail } = useToolRunner();

  const submit = () => {
    if (!files.length) return fail("Please choose at least one .xlsx file.");
    const form = new FormData();
    files.forEach((f) => form.append("files", f));
    run("/tools/report-merger", form);
  };

  return (
    <ToolPage
      title="Report Merger"
      description="Merge several “Report Table” exports into one workbook. The title row above the header is removed automatically. Sheet1 has all data with a blank row between files; Sheet2 lists every IMEI as IMEI1, IMEI2 …"
    >
      <Card>
        <FileDrop
          multiple
          label="Report files"
          files={files}
          onChange={setFiles}
          hint="Files are merged in the order shown"
        />
        <RunButton loading={loading} disabled={!files.length} onClick={submit}>
          Merge {files.length > 1 ? `${files.length} files` : "files"}
        </RunButton>
      </Card>
      <ResultPanel result={result} error={error} />
    </ToolPage>
  );
}
