"use client";

import { useCallback, useState } from "react";
import { postForDownload, saveBlob, ToolResult } from "./api";

/** Runs one backend tool, downloads the file it returns and keeps the summary. */
export function useToolRunner() {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<ToolResult | null>(null);

  const run = useCallback(async (path: string, form: FormData) => {
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const res = await postForDownload(path, form);
      saveBlob(res.blob, res.filename);
      setResult(res);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Something went wrong.");
    } finally {
      setLoading(false);
    }
  }, []);

  const fail = useCallback((message: string) => {
    setResult(null);
    setError(message);
  }, []);

  return { loading, error, result, run, fail };
}
