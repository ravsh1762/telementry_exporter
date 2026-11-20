"use client";

import { useSyncExternalStore } from "react";
import { applyTheme, DEFAULT_THEME, Theme, THEMES } from "@/lib/theme";

function subscribe(cb: () => void) {
  window.addEventListener("themechange", cb);
  window.addEventListener("storage", cb);
  return () => {
    window.removeEventListener("themechange", cb);
    window.removeEventListener("storage", cb);
  };
}

const getSnapshot = () =>
  (document.documentElement.dataset.theme as Theme | undefined) ??
  DEFAULT_THEME;
const getServerSnapshot = () => null;

const ICONS: Record<Theme, React.ReactNode> = {
  light: (
    <svg width="13" height="13" viewBox="0 0 16 16" fill="none" aria-hidden>
      <circle cx="8" cy="8" r="3" stroke="currentColor" strokeWidth="1.4" />
      <path
        d="M8 1.5v1.8M8 12.7v1.8M1.5 8h1.8M12.7 8h1.8M3.4 3.4l1.3 1.3M11.3 11.3l1.3 1.3M12.6 3.4l-1.3 1.3M4.7 11.3l-1.3 1.3"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
      />
    </svg>
  ),
  dark: (
    <svg width="13" height="13" viewBox="0 0 16 16" fill="none" aria-hidden>
      <path
        d="M13.5 9.6A5.8 5.8 0 0 1 6.4 2.5a5.8 5.8 0 1 0 7.1 7.1Z"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinejoin="round"
      />
    </svg>
  ),
  teal: (
    <svg width="13" height="13" viewBox="0 0 16 16" fill="none" aria-hidden>
      <circle cx="8" cy="8" r="5.8" stroke="currentColor" strokeWidth="1.4" />
      <path d="M8 2.2a5.8 5.8 0 0 0 0 11.6Z" fill="currentColor" />
    </svg>
  ),
};

/** Light / Dark Mode / Dark Theme switch. Choice is saved in localStorage. */
export function ThemeToggle() {
  const theme = useSyncExternalStore(subscribe, getSnapshot, getServerSnapshot);

  return (
    <div
      role="radiogroup"
      aria-label="Colour theme"
      className="inline-flex items-center gap-0.5 rounded-full border border-slate-200 bg-white p-0.5 h-8"
    >
      {THEMES.map((t) => {
        const active = theme === t.id;
        return (
          <button
            key={t.id}
            type="button"
            role="radio"
            aria-checked={active}
            title={t.title}
            onClick={() => applyTheme(t.id)}
            className={`inline-flex items-center gap-1.5 rounded-full px-2.5 h-7 text-xs font-semibold transition-all cursor-pointer ${
              active
                ? "bg-blue-600 text-white shadow-sm"
                : "text-slate-500 hover:text-slate-800"
            }`}
          >
            {ICONS[t.id]}
            <span className="hidden sm:inline">{t.label}</span>
          </button>
        );
      })}
    </div>
  );
}
