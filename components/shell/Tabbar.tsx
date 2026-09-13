"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import Icon from "./Icon";
import { tabItems, isActive } from "./nav-config";
import type { ShellUser, ShellStats } from "./types";

export default function Tabbar({ user, stats, onMenu }: { user: ShellUser; stats: ShellStats | null; onMenu: () => void }) {
  const pathname = usePathname() || "/";
  const items = tabItems(user);
  return (
    <nav className="tabbar no-print" aria-label="Hızlı gezinme">
      {items.map((it) => {
        const act = isActive(it, pathname);
        const primary = it.href === "/panel";
        const badge = it.badge === "acik" && stats ? stats.acik : 0;
        return (
          <Link key={it.href} href={it.href} className={`tab-item ${act ? "active" : ""} ${primary ? "primary" : ""}`} aria-current={act ? "page" : undefined}>
            <span className="tab-ico"><Icon name={it.icon} size={primary ? 22 : 21} /></span>
            <span>{it.label}</span>
            {badge > 0 && <span className="badge count">{badge > 99 ? "99+" : badge}</span>}
          </Link>
        );
      })}
      <button type="button" className="tab-item" onClick={onMenu} aria-label="Menü">
        <span className="tab-ico"><Icon name="menu" size={21} /></span>
        <span>Menü</span>
      </button>
    </nav>
  );
}
