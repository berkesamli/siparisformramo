"use client";

/* eslint-disable @next/next/no-img-element */
import { useState } from "react";
import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";

export interface NavUser {
  kind: "admin" | "dealer";
  name: string;
}

export default function NavBar({ user }: { user: NavUser }) {
  const [open, setOpen] = useState(false);
  const router = useRouter();
  const pathname = usePathname();

  const links: { href: string; label: string }[] =
    user.kind === "admin"
      ? [{ href: "/yonetim", label: "Bayiler" }]
      : [
          { href: "/panel", label: "Özet" },
          { href: "/panel/cerceve", label: "Online Çerçeve" },
          { href: "/panel/siparisler", label: "Siparişler" },
          { href: "/panel/ayarlar", label: "Fiyat & Ayarlar" },
        ];

  async function logout() {
    setOpen(false);
    await fetch("/api/auth/logout", { method: "POST" });
    router.push("/giris");
    router.refresh();
  }

  const roleLabel = user.kind === "admin" ? "Olga Yönetici" : "Bayi";

  return (
    <>
      <nav className="topnav">
        <button className="hamburger" aria-label="Menüyü aç" onClick={() => setOpen(true)}>
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round">
            <line x1="3" y1="6" x2="21" y2="6" />
            <line x1="3" y1="12" x2="21" y2="12" />
            <line x1="3" y1="18" x2="21" y2="18" />
          </svg>
        </button>

        <Link href={user.kind === "admin" ? "/yonetim" : "/panel"} className="logo">
          <img src="/logo.png" alt="Olga Çerçeve" style={{ height: 32, width: "auto", display: "block" }} />
          <span className="logo-sub" style={{ marginTop: 3 }}>BAYİ</span>
        </Link>

        <div className="nav-links">
          {links.map((l) => (
            <Link key={l.href} href={l.href} className="navlink">
              {l.label}
            </Link>
          ))}
        </div>

        <span className="spacer" />
        <span className="nav-user-area">
          <span className="user">
            {user.name} ({roleLabel})
          </span>
          <button className="btn small secondary" onClick={logout}>
            Çıkış
          </button>
        </span>
      </nav>

      {open && (
        <>
          <div className="drawer-backdrop" onClick={() => setOpen(false)} />
          <aside className="drawer">
            <div style={{ display: "flex", alignItems: "center", marginBottom: 22 }}>
              <img src="/logo.png" alt="Olga Çerçeve" style={{ height: 30, width: "auto" }} />
              <span style={{ flex: 1 }} />
              <button className="hamburger" style={{ display: "flex" }} aria-label="Menüyü kapat" onClick={() => setOpen(false)}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round">
                  <line x1="5" y1="5" x2="19" y2="19" />
                  <line x1="19" y1="5" x2="5" y2="19" />
                </svg>
              </button>
            </div>
            {links.map((l) => (
              <Link
                key={l.href}
                href={l.href}
                className={`drawer-link ${pathname === l.href ? "active" : ""}`}
                onClick={() => setOpen(false)}
              >
                {l.label}
              </Link>
            ))}
            <div style={{ flex: 1 }} />
            <div style={{ borderTop: "1px solid var(--border)", paddingTop: 14 }}>
              <div style={{ fontSize: 13, color: "var(--text-2)", marginBottom: 10 }}>
                {user.name}
                <br />
                <span style={{ color: "var(--muted)", fontSize: 12 }}>{roleLabel}</span>
              </div>
              <button className="btn small secondary" style={{ width: "100%" }} onClick={logout}>
                Çıkış Yap
              </button>
            </div>
          </aside>
        </>
      )}
    </>
  );
}
