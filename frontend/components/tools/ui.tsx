"use client";

import React from "react";
import { ToolResult } from "@/lib/api";
import { saveBlob } from "@/lib/api";
import { SiteHeader } from "../SiteHeader";

export const inputClass =
  "w-full bg-white border border-slate-200 rounded-lg px-3 py-2.5 text-sm text-gray-900 outline-none transition-all hover:border-slate-400 focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500";

export function ToolPage({
  title,
  description,
  children,
}: {
  title: string;
  description: React.ReactNode;
  children: React.ReactNode;
}) {
  return (
    <div className="min-h-screen bg-slate-100 dot-bg">
      <SiteHeader subtitle="Excel Tools" />
      <main className="max-w-[1600px] mx-auto px-6 py-8 space-y-6">
        <div>
          <h1 className="text-2xl font-bold text-slate-800 tracking-tight">
            {title}
          </h1>
          <p className="text-sm text-slate-500 mt-1">{description}</p>
        </div>
        {children}
      </main>
    </div>
  );
}

export function Card({
  children,
  className = "",
}: {
  children: React.ReactNode;
  className?: string;
}) {
  return (
    <div
      className={`bg-white border border-slate-200 rounded-xl p-5 shadow-sm space-y-5 ${className}`}
    >
      {children}
    </div>
  );
}

export function Field({
  label,
  hint,
  children,
}: {
  label: string;
  hint?: string;
  children: React.ReactNode;
}) {
  return (
    <label className="flex flex-col gap-1.5">
      <span className="text-xs font-semibold tracking-widest uppercase text-slate-400">
        {label}
      </span>
      {children}
      {hint && <span className="text-xs text-slate-400">{hint}</span>}
    </label>
  );
}

export function Checkbox({
  checked,
  onChange,
  label,
}: {
  checked: boolean;
  onChange: (v: boolean) => void;
  label: string;
}) {
  return (
    <label className="flex items-center gap-2 text-sm text-slate-600 cursor-pointer">
      <input
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.target.checked)}
        className="w-4 h-4 accent-blue-600"
      />
      {label}
    </label>
  );
}

export function RunButton({
  loading,
  disabled,
  onClick,
  children,
}: {
  loading: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={loading || disabled}
      className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold text-sm rounded-lg px-4 py-3 shadow-sm transition-all disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center gap-2"
    >
      {loading && (
        <span className="w-4 h-4 border-2 border-white/40 border-t-white rounded-full animate-spin-slow" />
      )}
      {loading ? "Processing…" : children}
    </button>
  );
}

export function ResultPanel({
  result,
  error,
}: {
  result: ToolResult | null;
  error: string | null;
}) {
  if (error) {
    return (
      <div className="bg-red-50 border border-red-200 text-red-700 text-sm rounded-xl px-4 py-3 animate-fade-in">
        {error}
      </div>
    );
  }
  if (!result) return null;

  const stats = Object.entries(result.summary?.stats ?? {});
  const notes = result.summary?.notes ?? [];

  return (
    <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-5 space-y-4 animate-slide-up">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <p className="text-sm font-semibold text-emerald-800">
            Done — {result.filename} downloaded
          </p>
        </div>
        <button
          type="button"
          onClick={() => saveBlob(result.blob, result.filename)}
          className="text-xs font-semibold text-emerald-700 border border-emerald-300 bg-white hover:bg-emerald-100 rounded-lg px-3 py-1.5"
        >
          Download again
        </button>
      </div>

      {stats.length > 0 && (
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {stats.map(([label, value]) => (
            <div
              key={label}
              className="bg-white border border-emerald-100 rounded-lg px-3 py-2"
            >
              <p className="text-[10px] font-semibold tracking-wider uppercase text-emerald-600">
                {label}
              </p>
              <p className="text-lg font-bold font-mono text-emerald-800">
                {typeof value === "number" ? value.toLocaleString("en-IN") : value}
              </p>
            </div>
          ))}
        </div>
      )}

      {notes.length > 0 && (
        <ul className="text-xs text-emerald-800 space-y-0.5 font-mono">
          {notes.map((n, i) => (
            <li key={i}>{n}</li>
          ))}
        </ul>
      )}
    </div>
  );
}
