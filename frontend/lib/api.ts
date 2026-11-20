"use client";

/** Base URL of the Express backend. Override with NEXT_PUBLIC_API_URL if needed. */
export const API_BASE = process.env.NEXT_PUBLIC_API_URL || "http://localhost:4000";

export interface ToolSummary {
  stats?: Record<string, string | number>;
  notes?: string[];
}

export interface ToolResult {
  blob: Blob;
  filename: string;
  summary: ToolSummary | null;
}

export function saveBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** POSTs a form to a backend tool and returns the file it sends back. */
export async function postForDownload(
  path: string,
  form: FormData,
): Promise<ToolResult> {
  let response: Response;
  try {
    response = await fetch(`${API_BASE}${path}`, { method: "POST", body: form });
  } catch {
    throw new Error(
      `Cannot reach the backend at ${API_BASE}. Make sure it is running (npm run dev).`,
    );
  }

  if (!response.ok) {
    let message = `Request failed (${response.status})`;
    try {
      const data = await response.json();
      if (data?.error) message = data.error;
    } catch {
      /* not JSON - keep the default message */
    }
    throw new Error(message);
  }

  const disposition = response.headers.get("Content-Disposition") || "";
  const match = /filename="?([^";]+)"?/i.exec(disposition);
  const filename = match ? decodeURIComponent(match[1]) : "result.xlsx";

  let summary: ToolSummary | null = null;
  const raw = response.headers.get("X-Tool-Summary");
  if (raw) {
    try {
      summary = JSON.parse(decodeURIComponent(raw));
    } catch {
      summary = null;
    }
  }

  return { blob: await response.blob(), filename, summary };
}
