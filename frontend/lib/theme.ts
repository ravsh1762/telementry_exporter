// Single source of truth for the theme system.
// The colours themselves live in app/globals.css ([data-theme="dark"] / [data-theme="teal"]).

export type Theme = "light" | "dark" | "teal";

export const THEME_STORAGE_KEY = "theme";

/** Theme shown the first time someone opens the site (before they pick one). */
export const DEFAULT_THEME: Theme = "dark";

export const THEMES: { id: Theme; label: string; title: string }[] = [
  { id: "light", label: "Light", title: "Light theme" },
  { id: "dark", label: "Dark", title: "Dark Mode – near-black with cyan glow" },
  { id: "teal", label: "Teal", title: "Dark Theme – soft teal-slate" },
];

/**
 * Runs in <head> BEFORE the page paints, so there is no white flash on reload.
 * Reads the saved choice, falls back to DEFAULT_THEME.
 */
export const THEME_INIT_SCRIPT = `(function(){var d=document.documentElement,t;try{t=localStorage.getItem(${JSON.stringify(
  THEME_STORAGE_KEY,
)})}catch(e){}if(t!=="light"&&t!=="dark"&&t!=="teal")t=${JSON.stringify(
  DEFAULT_THEME,
)};d.dataset.theme=t;d.classList.toggle("dark",t!=="light")})();`;

export function applyTheme(theme: Theme) {
  const root = document.documentElement;
  root.dataset.theme = theme;
  root.classList.toggle("dark", theme !== "light");
  try {
    localStorage.setItem(THEME_STORAGE_KEY, theme);
  } catch {
    /* private mode – the choice just won't persist */
  }
  window.dispatchEvent(new Event("themechange"));
}
