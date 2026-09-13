"use client";

import { useEffect, useState } from "react";
import Icon from "./Icon";

export const THEME_KEY = "olga-theme";

export function applyTheme(t: "dark" | "light") {
  document.documentElement.setAttribute("data-theme", t);
  try { localStorage.setItem(THEME_KEY, t); } catch { /* özel pencere */ }
}

export default function ThemeToggle({ label = false, className = "" }: { label?: boolean; className?: string }) {
  const [theme, setTheme] = useState<"dark" | "light">("dark");

  useEffect(() => {
    const cur = document.documentElement.getAttribute("data-theme");
    setTheme(cur === "light" ? "light" : "dark");
  }, []);

  function toggle() {
    const next = theme === "dark" ? "light" : "dark";
    setTheme(next);
    applyTheme(next);
  }

  const dark = theme === "dark";
  return (
    <button
      type="button"
      className={`btn ${label ? "secondary small" : "icon ghost"} theme-btn ${className}`}
      onClick={toggle}
      title={dark ? "Açık temaya geç" : "Koyu temaya geç"}
      aria-label={dark ? "Açık temaya geç" : "Koyu temaya geç"}
    >
      <Icon name={dark ? "sun" : "moon"} size={18} />
      {label && <span className="btn-label">{dark ? "Açık Tema" : "Koyu Tema"}</span>}
    </button>
  );
}
