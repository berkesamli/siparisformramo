"use client";

// Uygulama kabuğu: sol kenar çubuğu + üst çubuk + içerik + telefonda alt sekme
// çubuğu + genel arama paleti. Sunucu sayfaları children olarak gelir.

import { useCallback, useEffect, useState, type ReactNode } from "react";
import { usePathname, useRouter } from "next/navigation";
import Sidebar from "./Sidebar";
import Topbar from "./Topbar";
import Tabbar from "./Tabbar";
import SearchPalette from "./SearchPalette";
import AnnouncementModal from "./AnnouncementModal";
import { STATS_KEY, STATS_TTL_MS, type Duyuru, type ShellStats, type ShellUser } from "./types";

const SB_KEY = "olga-sb";
const BOS_DUYURU: Duyuru[] = []; // sabit referans — her render'da yeni dizi üretmesin

export default function AppShell({
  user,
  duyurular = BOS_DUYURU,
  children,
}: {
  user: ShellUser;
  duyurular?: Duyuru[];
  children: ReactNode;
}) {
  const pathname = usePathname();
  const router = useRouter();
  const [drawer, setDrawer] = useState(false);
  const [search, setSearch] = useState(false);
  const [rail, setRail] = useState(false);
  const [stats, setStats] = useState<ShellStats | null>(null);

  // Ray tercihi (html[data-sb]) — bayrağı giriş betiği boyamadan önce koyar
  useEffect(() => {
    setRail(document.documentElement.getAttribute("data-sb") === "rail");
  }, []);

  // Sayfa değişince çekmece kapanır
  useEffect(() => { setDrawer(false); }, [pathname]);

  // Çekmece açıkken arka plan kaymasın (telefon)
  useEffect(() => {
    if (!drawer) return;
    const prev = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { document.body.style.overflow = prev; };
  }, [drawer]);

  // Klavye: Ctrl/⌘+K arama, Esc çekmece
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        setSearch((s) => !s);
      } else if (e.key === "Escape") {
        setDrawer(false);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // Hafif sayaçlar: oturum önbelleği (3 dk) — her sayfada Blob okunmasın
  const loadStats = useCallback(async (force = false) => {
    if (user.role !== "staff") return;
    if (!force) {
      try {
        const raw = sessionStorage.getItem(STATS_KEY);
        if (raw) {
          const c = JSON.parse(raw) as { ts: number; data: ShellStats };
          if (Date.now() - c.ts < STATS_TTL_MS) { setStats(c.data); return; }
          setStats(c.data); // eskisini göster, arkada yenile
        }
      } catch { /* yok */ }
    }
    try {
      const r = await fetch("/api/dashboard?lite=1");
      const d = await r.json();
      if (d?.ok && d.lite) {
        setStats(d.lite as ShellStats);
        try { sessionStorage.setItem(STATS_KEY, JSON.stringify({ ts: Date.now(), data: d.lite })); } catch { /* yok */ }
      }
    } catch { /* sessiz */ }
  }, [user.role]);

  useEffect(() => { loadStats(); }, [loadStats, pathname]);

  const toggleRail = useCallback(() => {
    const next = !rail;
    setRail(next);
    if (next) document.documentElement.setAttribute("data-sb", "rail");
    else document.documentElement.removeAttribute("data-sb");
    try { localStorage.setItem(SB_KEY, next ? "rail" : "full"); } catch { /* yok */ }
  }, [rail]);

  const logout = useCallback(async () => {
    setDrawer(false);
    try { sessionStorage.removeItem(STATS_KEY); } catch { /* yok */ }
    await fetch("/api/auth/logout", { method: "POST" });
    router.push("/giris");
    router.refresh();
  }, [router]);

  return (
    <div className="app">
      {drawer && <div className="sb-backdrop" onClick={() => setDrawer(false)} aria-hidden />}
      <Sidebar
        user={user}
        stats={stats}
        open={drawer}
        rail={rail}
        onClose={() => setDrawer(false)}
        onToggleRail={toggleRail}
        onLogout={logout}
      />
      <div className="app-main">
        <Topbar
          user={user}
          onMenu={() => setDrawer(true)}
          onSearch={() => setSearch(true)}
          onLogout={logout}
          onStatsDirty={() => loadStats(true)}
          duyurular={duyurular}
        />
        <div className="app-content">{children}</div>
      </div>
      <Tabbar user={user} stats={stats} onMenu={() => setDrawer(true)} onSearch={() => setSearch(true)} />
      <AnnouncementModal duyurular={duyurular} username={user.username} />
      {search && <SearchPalette user={user} onClose={() => setSearch(false)} />}
    </div>
  );
}
