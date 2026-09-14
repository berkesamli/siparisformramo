"use client";

import { useEffect, useState } from "react";
import Icon from "./Icon";

export const THEME_KEY = "olga-theme";

export function applyTheme(t: "dark" | "light") {
  document.documentElement.setAttribute("data-theme", t);
  try { localStorage.setItem(THEME_KEY, t); } catch { /* özel pencere */ }
}

export default function ThemeToggle({ label = false, className = "" }: { label?: boolean; className?: string }) {
  const [theme, setTheme] = useState<"dark" | "light">("light");

  // Birden fazla anahtar (üst çubuk + kenar çubuğu) aynı anda var: biri
  // değiştirince diğeri html[data-theme] üzerinden senkron kalır.
  useEffect(() => {
    const oku = () => {
      const cur = document.documentElement.getAttribute("data-theme");
      setTheme(cur === "dark" ? "dark" : "light");
    };
    oku();
    const mo = new MutationObserver(oku);
    mo.observe(document.documentElement, { attributes: true, attributeFilter: ["data-theme"] });
    return () => mo.disconnect();
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
