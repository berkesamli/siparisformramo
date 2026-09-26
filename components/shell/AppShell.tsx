"use client";

// Uygulama kabuğu: sol kenar çubuğu + üst çubuk + içerik + telefonda alt sekme
// çubuğu + genel arama paleti. Sunucu sayfaları children olarak gelir.

import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
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

  // Yüzen kutu düzeninde aktif menü "dili": kenar çubuğu ile içerik kutusu
  // arasındaki 16 px boşluğa, aktif öğenin hizasına içerik renginde köprü
  // (+ içbükey köşeler) çizilir. Menü kaydırılır; öğe görünür alandan çıkınca
  // dil gizlenir. Ölçüm rAF ile birleştirilir: rota, ray, kaydırma, boyut,
  // rozet/alt menü değişimleri tetikler.
  const dilRef = useRef<HTMLSpanElement>(null);
  useEffect(() => {
    const el = dilRef.current;
    if (!el) return;
    let raf = 0;
    const olc = () => {
      raf = 0;
      const sb = document.querySelector<HTMLElement>(".sb");
      const nav = sb?.querySelector<HTMLElement>(".sb-nav");
      const act = nav?.querySelector<HTMLElement>(".sb-link.active");
      const kutu = window.matchMedia("(min-width: 1181px)").matches;
      if (!sb || !nav || !act || !kutu) { el.classList.remove("on"); return; }
      const s = sb.getBoundingClientRect();
      const n = nav.getBoundingClientRect();
      const a = act.getBoundingClientRect();
      if (a.top < n.top - 1 || a.bottom > n.bottom - 26) { el.classList.remove("on"); return; }
      el.style.left = `${Math.round(s.right)}px`;
      el.style.top = `${Math.round(a.top - 16)}px`;
      el.style.height = `${Math.round(a.height + 32)}px`;
      el.classList.add("on");
    };
    const iste = () => { if (!raf) raf = requestAnimationFrame(olc); };
    iste();
    const sb = document.querySelector<HTMLElement>(".sb");
    const nav = sb?.querySelector<HTMLElement>(".sb-nav");
    nav?.addEventListener("scroll", iste, { passive: true });
    window.addEventListener("resize", iste);
    const ro = new ResizeObserver(iste);
    if (sb) ro.observe(sb);
    const mo = new MutationObserver(iste);
    if (sb) mo.observe(sb, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ["class"] });
    // Ray geçişi 0.22 sn animasyonlu: bitince tekrar ölç
    const t = [60, 260, 420].map((ms) => setTimeout(iste, ms));
    return () => {
      cancelAnimationFrame(raf);
      nav?.removeEventListener("scroll", iste);
      window.removeEventListener("resize", iste);
      ro.disconnect();
      mo.disconnect();
      t.forEach(clearTimeout);
    };
  }, [pathname, rail]);

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
      <span ref={dilRef} className="sb-dil" aria-hidden="true" />
      <Tabbar user={user} stats={stats} onMenu={() => setDrawer(true)} onSearch={() => setSearch(true)} />
      <AnnouncementModal duyurular={duyurular} username={user.username} />
      {search && <SearchPalette user={user} onClose={() => setSearch(false)} />}
    </div>
  );
}
