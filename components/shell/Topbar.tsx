"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import Icon from "./Icon";
import ThemeToggle from "./ThemeToggle";
import UserMenu from "./UserMenu";
import NotificationBell from "./NotificationBell";
import { breadcrumbsFor } from "./nav-config";
import type { Duyuru, ShellUser } from "./types";

const BOS_DUYURU: Duyuru[] = []; // sabit referans — zilin effect bağımlılığı her render'da değişmesin

export default function Topbar({
  user,
  onMenu,
  onSearch,
  onLogout,
  onStatsDirty,
  duyurular = BOS_DUYURU,
}: {
  user: ShellUser;
  onMenu: () => void;
  onSearch: () => void;
  onLogout: () => void;
  onStatsDirty: () => void;
  duyurular?: Duyuru[];
}) {
  const pathname = usePathname() || "/";
  const crumbs = breadcrumbsFor(user, pathname);
  const cur = crumbs[crumbs.length - 1];

  return (
    <header className="topbar">
      <button type="button" className="btn icon ghost tb-burger" onClick={onMenu} aria-label="Menüyü aç">
        <Icon name="menu" size={20} />
      </button>

      <nav className="tb-crumbs" aria-label="Sayfa yolu">
        {crumbs.map((c, i) => {
          const last = i === crumbs.length - 1;
          return (
            <span key={c.href} style={{ display: "contents" }}>
              {i > 0 && <span className="sep">›</span>}
              {last ? <span className="cur">{c.label}</span> : <Link href={c.href}>{c.label}</Link>}
            </span>
          );
        })}
      </nav>
      <span className="tb-title-mobile">{cur?.label}</span>

      <button type="button" className="tb-search" onClick={onSearch} aria-label="Ara" title="Ara (Ctrl+K)">
        <Icon name="search" size={17} />
        <span className="tb-search-text">Sipariş, müşteri, ürün ara…</span>
        <span className="kbd">Ctrl K</span>
      </button>

      <div className="tb-actions">
        {user.role === "staff" && <NotificationBell onStatsDirty={onStatsDirty} duyurular={duyurular} username={user.username} />}
        <ThemeToggle />
        <UserMenu user={user} onLogout={onLogout} />
      </div>
    </header>
  );
}
