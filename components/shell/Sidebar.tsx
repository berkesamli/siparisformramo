"use client";

/* eslint-disable @next/next/no-img-element */

import { useEffect, useRef, useState } from "react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import Icon from "./Icon";
import ThemeToggle from "./ThemeToggle";
import { navGroups, activeHref, initials, isActive, type NavItem } from "./nav-config";
import type { ShellUser, ShellStats } from "./types";

const fmtK = (n: number) => (n >= 10000 ? `${(n / 1000).toFixed(1)}k` : n.toLocaleString("tr-TR"));

export default function Sidebar({
  user,
  stats,
  open,
  rail,
  onClose,
  onToggleRail,
  onLogout,
}: {
  user: ShellUser;
  stats: ShellStats | null;
  open: boolean;
  rail: boolean;
  onClose: () => void;
  onToggleRail: () => void;
  onLogout: () => void;
}) {
  const pathname = usePathname() || "/";
  const groups = navGroups(user);
  const active = activeHref(user, pathname);
  const [openGroups, setOpenGroups] = useState<Record<string, boolean>>({});

  // Aktif alt öğesi olan açılır menü kendiliğinden açık gelsin
  useEffect(() => {
    const next: Record<string, boolean> = {};
    for (const g of groups) {
      for (const it of g.items) {
        if (it.children && it.children.some((c) => isActive(c, pathname))) next[it.href] = true;
      }
    }
    setOpenGroups((o) => ({ ...o, ...next }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pathname]);

  // Geçerli sayfanın menü öğesi .sb-foot / alt solma maskesinin arkasında yarım kalmasın:
  // rota değişince yalnızca .sb-nav kaydırılır (sayfa değil), öğe zaten görünürse dokunulmaz.
  // Aktif öğe kapalı bir alt menüdeyse ilk turda DOM'da yoktur; openGroups açılınca tekrar denenir,
  // ancak her rota için yalnızca bir kez kaydırılır (elle grup açıp kapatmayla çakışmasın).
  const navRef = useRef<HTMLElement>(null);
  const scrolledFor = useRef<string | null>(null);
  useEffect(() => {
    if (!active || scrolledFor.current === active) return;
    const nav = navRef.current;
    const el = nav?.querySelector<HTMLElement>(".sb-link.active");
    if (!nav || !el) return;
    scrolledFor.current = active;
    const pad = 32; // .sb-nav alt solma maskesi (28px) + pay
    const n = nav.getBoundingClientRect();
    const r = el.getBoundingClientRect();
    if (r.bottom > n.bottom - pad) nav.scrollTop += r.bottom - (n.bottom - pad);
    else if (r.top < n.top + pad) nav.scrollTop -= n.top + pad - r.top;
  }, [active, openGroups]);

  const badgeValue = (it: NavItem): number => {
    if (!it.badge || !stats) return 0;
    if (it.badge === "acik") return stats.acik;
    if (it.badge === "perakendeAcik") return stats.perakendeAcik;
    if (it.badge === "kontrolsuz") return stats.kontrolsuz;
    if (it.badge === "mesaj") return stats.mesajOkunmamis || 0;
    return 0;
  };

  const roleLabel = user.role === "staff" ? "Çalışan" : "Müşteri / Bayi";
  const delta = stats ? stats.bugun - stats.dun : 0;

  function renderItem(it: NavItem) {
    const act = active === it.href;
    const badge = badgeValue(it);
    const hasKids = !!it.children?.length;
    const isOpen = hasKids && (openGroups[it.href] ?? false);
    return (
      <div key={it.href}>
        <div style={{ display: "flex", alignItems: "stretch", gap: 2 }}>
          <Link
            href={it.href}
            className={`sb-link ${act ? "active" : ""} ${isOpen ? "open" : ""}`}
            title={rail ? it.label : undefined}
            onClick={onClose}
            aria-current={act ? "page" : undefined}
            style={{ flex: 1 }}
          >
            <span className="sb-link-icon"><Icon name={it.icon} size={19} /></span>
            <span className="sb-link-label">{it.label}</span>
            {badge > 0 && <span className="badge count" aria-label={`${badge} kayıt`}>{badge > 99 ? "99+" : badge}</span>}
            {hasKids && (
              <span
                className="sb-link-chev"
                role="button"
                aria-label={isOpen ? "Alt menüyü kapat" : "Alt menüyü aç"}
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  setOpenGroups((o) => ({ ...o, [it.href]: !isOpen }));
                }}
              >
                <Icon name="chevron-down" size={16} />
              </span>
            )}
          </Link>
        </div>
        {hasKids && (isOpen || rail) && (
          <div className="sb-sub">
            {it.children!.map((c) => renderItem(c))}
          </div>
        )}
      </div>
    );
  }

  return (
    <aside className={`sb ${open ? "open" : ""}`} aria-label="Ana menü">
      <div className="sb-brand">
        <Link href="/" className="sb-brand-link" onClick={onClose} aria-label="Olga Çerçeve — Ana Sayfa">
          <span className="sb-brand-mark"><img src="/logo.png" alt="" /></span>
          <span className="sb-brand-text">
            <span className="sb-brand-word">olga</span>
            <span className="sb-brand-sub">ÇERÇEVE</span>
          </span>
        </Link>
        <button type="button" className="btn icon ghost small sb-close" onClick={onClose} aria-label="Menüyü kapat">
          <Icon name="x" size={18} />
        </button>
      </div>

      {/* Daralt / genişlet tutamacı — kenar çubuğunun sağ kenarında, ray modunda da görünür */}
      <button
        type="button"
        className="sb-handle"
        onClick={onToggleRail}
        title={rail ? "Menüyü genişlet" : "Menüyü daralt"}
        aria-label={rail ? "Menüyü genişlet" : "Menüyü daralt"}
        aria-expanded={!rail}
      >
        <Icon name={rail ? "chevron-right" : "chevron-left"} size={15} strokeWidth={2.4} />
      </button>

      <div className="sb-user">
        <span className="avatar lg" aria-hidden>{initials(user.name)}</span>
        <div className="sb-user-name" title={user.name}>{user.name}</div>
        <div className="sb-user-role">
          <span>{roleLabel}</span>
          {user.owner && <span className="badge brand">Sahip</span>}
          {!user.owner && user.finance && <span className="badge info">Finans</span>}
        </div>
        {user.role === "staff" && (
          <div className="sb-stats" aria-label="Bugünün özeti">
            <Link href="/panel/siparisler" className="sb-stat" title="Bugün girilen siparişler">
              <span className="sb-stat-icon"><Icon name="inbox" size={16} /></span>
              <span className="sb-stat-val">{stats ? fmtK(stats.bugun) : "—"}</span>
              <span className="sb-stat-lbl">Bugün</span>
              {stats && delta !== 0 && (
                <span className={`sb-stat-delta ${delta < 0 ? "down" : ""}`}>{delta > 0 ? "+" : ""}{delta}</span>
              )}
            </Link>
            <Link href="/panel/siparisler" className="sb-stat" title="Açık (tamamlanmamış) siparişler">
              <span className="sb-stat-icon"><Icon name="activity" size={16} /></span>
              <span className="sb-stat-val">{stats ? fmtK(stats.acik + stats.perakendeAcik) : "—"}</span>
              <span className="sb-stat-lbl">Açık</span>
            </Link>
            <Link href="/panel/siparisler" className="sb-stat" title="Bu ay girilen siparişler">
              <span className="sb-stat-icon"><Icon name="calendar" size={16} /></span>
              <span className="sb-stat-val">{stats ? fmtK(stats.ayAdet) : "—"}</span>
              <span className="sb-stat-lbl">Bu Ay</span>
            </Link>
          </div>
        )}
      </div>

      <nav className="sb-nav" ref={navRef}>
        {groups.map((g) => (
          <div className="sb-group" key={g.title}>
            <div className="sb-group-title">{g.title}</div>
            {g.items.map((it) => renderItem(it))}
          </div>
        ))}
      </nav>

      <div className="sb-foot">
        <div className="sb-foot-row">
          <ThemeToggle label />
          <span className="spacer" />
          <button type="button" className="btn secondary small" onClick={onLogout} title="Çıkış yap">
            <Icon name="log-out" size={16} />
            <span className="btn-label">Çıkış</span>
          </button>
        </div>
        <div className="sb-foot-brand">
          <strong>OLGA ÇERÇEVE</strong>
          © {new Date().getFullYear()} · 0850 305 75 45
        </div>
      </div>
    </aside>
  );
}
