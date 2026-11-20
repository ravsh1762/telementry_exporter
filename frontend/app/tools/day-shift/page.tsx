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

export default function DayShiftPage() {
  const [files, setFiles] = useState<File[]>([]);
  const [days, setDays] = useState("1");
  const [keepTime, setKeepTime] = useState(false);
  const { loading, error, result, run, fail } = useToolRunner();

  const submit = () => {
    if (!files.length) return fail("Please choose an .xlsx file.");
    if (!/^-?\d+$/.test(days.trim())) return fail("Days must be a whole number.");
    const form = new FormData();
    form.append("file", files[0]);
    form.append("days", days.trim());
    form.append("keepTime", String(keepTime));
    run("/tools/day-shift", form);
  };

  return (
    <ToolPage
      title="Day Shift"
      description="Adds days to every TIMESTAMP. By default the result is the date only (YYYY-MM-DD)."
    >
      <Card>
        <FileDrop label="Excel file" files={files} onChange={setFiles} />
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 items-end">
          <Field label="Days to add" hint="Use a negative number to go back">
            <input
              type="number"
              value={days}
              onChange={(e) => setDays(e.target.value)}
              className={inputClass}
            />
          </Field>
          <Checkbox
            checked={keepTime}
            onChange={setKeepTime}
            label="Keep the time of day (YYYY-MM-DD HH:MM:SS)"
          />
        </div>
        <RunButton loading={loading} disabled={!files.length} onClick={submit}>
          Shift dates
        </RunButton>
      </Card>
      <ResultPanel result={result} error={error} />
    </ToolPage>
  );
}
