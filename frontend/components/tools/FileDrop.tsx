"use client";

import React, { useRef, useState } from "react";

const formatSize = (n: number) =>
  n < 1024
    ? `${n} B`
    : n < 1048576
      ? `${(n / 1024).toFixed(0)} KB`
      : `${(n / 1048576).toFixed(1)} MB`;

interface FileDropProps {
  files: File[];
  onChange: (files: File[]) => void;
  multiple?: boolean;
  label: string;
  hint?: string;
}

export function FileDrop({ files, onChange, multiple, label, hint }: FileDropProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  const [warning, setWarning] = useState("");

  const add = (incoming: FileList | File[]) => {
    const list = Array.from(incoming);
    const good = list.filter((f) => /\.xlsx$/i.test(f.name));
    setWarning(
      good.length < list.length ? "Only .xlsx files are supported — others were skipped." : "",
    );
    if (!good.length) return;
    if (!multiple) return onChange([good[0]]);
    const merged = [...files];
    for (const f of good) {
      if (!merged.some((x) => x.name === f.name && x.size === f.size)) merged.push(f);
    }
    onChange(merged);
  };

  return (
    <div className="flex flex-col gap-1.5">
      <span className="text-xs font-semibold tracking-widest uppercase text-slate-400">
        {label}
      </span>

      <div
        onClick={() => inputRef.current?.click()}
        onDragOver={(e) => {
          e.preventDefault();
          setOver(true);
        }}
        onDragLeave={() => setOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setOver(false);
          add(e.dataTransfer.files);
        }}
        className={`border-2 border-dashed rounded-lg px-4 py-6 text-center cursor-pointer text-sm transition-all ${
          over
            ? "border-blue-500 bg-blue-50 text-blue-600"
            : "border-slate-200 text-slate-500 hover:border-slate-400"
        }`}
      >
        <span className="font-semibold text-slate-700">
          Drop {multiple ? ".xlsx files" : "an .xlsx file"} here
        </span>{" "}
        or click to choose
        {hint && <div className="text-xs text-slate-400 mt-1">{hint}</div>}
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx"
          multiple={multiple}
          className="hidden"
          onChange={(e) => {
            if (e.target.files) add(e.target.files);
            e.target.value = "";
          }}
        />
      </div>

      {warning && <p className="text-xs text-amber-600">{warning}</p>}

      {files.length > 0 && (
        <ul className="divide-y divide-slate-100 border border-slate-200 rounded-lg">
          {files.map((f, i) => (
            <li key={`${f.name}-${i}`} className="flex items-center gap-3 px-3 py-2 text-sm">
              {multiple && <span className="text-slate-400 w-5 text-xs">{i + 1}</span>}
              <span className="flex-1 truncate text-slate-700">{f.name}</span>
              <span className="text-xs text-slate-400">{formatSize(f.size)}</span>
              <button
                type="button"
                title="Remove"
                onClick={() => onChange(files.filter((_, j) => j !== i))}
                className="text-slate-400 hover:text-red-500 px-1"
              >
                ×
              </button>
            </li>
          ))}
        </ul>
      )}

      {multiple && files.length > 1 && (
        <button
          type="button"
          onClick={() =>
            onChange(
              [...files].sort((a, b) =>
                a.name.localeCompare(b.name, undefined, { numeric: true }),
              ),
            )
          }
          className="self-start text-xs font-semibold text-blue-600 hover:underline"
        >
          Sort A→Z
        </button>
      )}
    </div>
  );
}
