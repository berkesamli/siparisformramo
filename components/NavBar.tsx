"use client";

/* eslint-disable @next/next/no-img-element */

import { useState } from "react";
import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";

interface NavUser {
  name: string;
  role: "staff" | "customer";
  owner?: boolean; // firma sahibi — raporları görebilir
  finance?: boolean; // finans ekranları (kasa, gider, çek/senet, personel)
}

export default function NavBar({ user }: { user: NavUser | null }) {
  const [open, setOpen] = useState(false);
  const router = useRouter();
  const pathname = usePathname();

  const links: { href: string; label: string }[] = [
    { href: "/", label: "Ana Sayfa" },
    { href: "/kataloglar", label: "Kataloglar" },
  ];
  if (user) {
    links.push({ href: "/portal", label: "Ürünler & Stok" });
    links.push({ href: "/portal/fiyat-listesi", label: "Toptan Fiyat Listesi" });
  }
  if (user?.role === "staff") {
    links.push({ href: "/panel", label: "Sipariş Paneli" });
    links.push({ href: "/panel/siparisler", label: "Siparişler" });
    links.push({ href: "/panel/perakende", label: "Perakende" });
    links.push({ href: "/etiket", label: "Etiket" });
    links.push({ href: "/panel/sms", label: "SMS" });
  }
  if (user?.role === "staff" && user.finance) {
    links.push({ href: "/panel/finans", label: "Finans" });
  }
  if (user?.role === "staff" && user.owner) {
    links.push({ href: "/panel/raporlar", label: "Raporlar" });
  }
  // Çekmece menüde Stok Yükle de listelensin
  const drawerLinks =
    user?.role === "staff"
      ? [...links, { href: "/panel/stok", label: "Stok Yükle" }]
      : links;

  async function logout() {
    setOpen(false);
    await fetch("/api/auth/logout", { method: "POST" });
    router.push("/giris");
    router.refresh();
  }

  return (
    <>
      <nav className="topnav">
        <button
          className="hamburger"
          aria-label="Menüyü aç"
          onClick={() => setOpen(true)}
        >
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round">
            <line x1="3" y1="6" x2="21" y2="6" />
            <line x1="3" y1="12" x2="21" y2="12" />
            <line x1="3" y1="18" x2="21" y2="18" />
          </svg>
        </button>

        <Link href="/" className="logo">
          <img src="/logo.png" alt="Olga Çerçeve" style={{ height: 32, width: "auto", display: "block" }} />
          <span className="logo-sub" style={{ marginTop: 3 }}>ÇERÇEVE</span>
        </Link>

        <div className="nav-links">
          {links
            .filter((l) => l.href !== "/")
            .map((l) => (
              <Link key={l.href} href={l.href} className="navlink">
                {l.label}
              </Link>
            ))}
        </div>

        <span className="spacer" />

        {user ? (
          <span className="nav-user-area">
            <span className="user">
              {user.name} ({user.role === "staff" ? "Çalışan" : "Müşteri"})
            </span>
            {user.role === "staff" && (
              <Link href="/panel/stok" className="btn small secondary">
                📦 Stok Yükle
              </Link>
            )}
            <button className="btn small secondary" onClick={logout}>
              Çıkış
            </button>
          </span>
        ) : (
          <Link href="/giris" className="btn small nav-login">
            Giriş Yap
          </Link>
        )}
      </nav>

      {/* Mobil yan menü */}
      {open && (
        <>
          <div className="drawer-backdrop" onClick={() => setOpen(false)} />
          <aside className="drawer">
            <div style={{ display: "flex", alignItems: "center", marginBottom: 22 }}>
              <img src="/logo.png" alt="Olga Çerçeve" style={{ height: 30, width: "auto" }} />
              <span style={{ flex: 1 }} />
              <button
                className="hamburger"
                style={{ display: "flex" }}
                aria-label="Menüyü kapat"
                onClick={() => setOpen(false)}
              >
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round">
                  <line x1="5" y1="5" x2="19" y2="19" />
                  <line x1="19" y1="5" x2="5" y2="19" />
                </svg>
              </button>
            </div>

            {drawerLinks.map((l) => (
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

            {user ? (
              <div style={{ borderTop: "1px solid var(--border)", paddingTop: 14 }}>
                <div style={{ fontSize: 13, color: "var(--text-2)", marginBottom: 10 }}>
                  {user.name}
                  <br />
                  <span style={{ color: "var(--muted)", fontSize: 12 }}>
                    {user.role === "staff" ? "Çalışan" : "Müşteri"}
                  </span>
                </div>
                <button className="btn small secondary" style={{ width: "100%" }} onClick={logout}>
                  Çıkış Yap
                </button>
              </div>
            ) : (
              <Link
                href="/giris"
                className="btn"
                style={{ justifyContent: "center" }}
                onClick={() => setOpen(false)}
              >
                Giriş Yap
              </Link>
            )}
          </aside>
        </>
      )}
    </>
  );
}
