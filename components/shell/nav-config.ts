// Menü, alt sekme çubuğu, kırıntı (breadcrumb) ve arama için TEK kaynak.
// Yetki bayrakları layout'ta sunucuda hesaplanır (data/users.ts) ve buraya gelir.

import type { IconName } from "./Icon";

export interface NavFlags {
  role: "staff" | "customer";
  owner?: boolean;    // firma sahibi
  finance?: boolean;  // finans ekranları (FINANS_AKTIF + yetki)
  maliyet?: boolean;  // alış fiyatları & kârlılık
  kur?: boolean;      // günlük kur belirleme
  raporlar?: boolean; // ciro/tahsilat raporları (finans yetkisi, FINANS_AKTIF'ten bağımsız)
}

export type BadgeKey = "acik" | "kontrolsuz" | "perakendeAcik";

export interface NavItem {
  href: string;
  label: string;
  icon: IconName;
  short?: string;        // alt sekme çubuğu için kısa ad
  badge?: BadgeKey;      // kenar çubuğunda sayaç rozeti
  children?: NavItem[];  // açılır alt menü
  exact?: boolean;       // yalnızca tam eşleşmede aktif
  keywords?: string[];   // arama paleti için ek anahtar kelimeler
}

export interface NavGroup {
  title: string;
  items: NavItem[];
}

export function navGroups(f: NavFlags): NavGroup[] {
  const staff = f.role === "staff";
  const groups: NavGroup[] = [];

  // "Genel" grubu müşterilerle paylaşılır; "Satışlarım" yalnızca çalışanlara eklenir.
  const genel: NavItem[] = [
    { href: "/", label: "Gösterge Paneli", icon: "home", short: "Panel", exact: true, keywords: ["ana sayfa", "dashboard", "özet"] },
  ];
  if (staff) {
    genel.push({ href: "/panel/satislarim", label: "Satışlarım", icon: "trending-up", keywords: ["cirom", "satışlarım", "kişisel ciro", "performans"] });
  }
  genel.push(
    { href: "/kataloglar", label: "Kataloglar", icon: "book", short: "Katalog", keywords: ["pdf", "dergi", "profil kataloğu", "teknik malzeme"] },
    { href: "/portal", label: "Ürünler & Stok", icon: "package", short: "Stok", exact: true, keywords: ["stok sorgula", "depo", "ankara", "istanbul"] },
    { href: "/portal/fiyat-listesi", label: "Toptan Fiyat Listesi", icon: "tag", short: "Fiyat", keywords: ["fiyat", "liste", "usd"] },
  );
  groups.push({ title: "Genel", items: genel });

  if (staff) {
    groups.push({
      title: "Toptan Sipariş",
      items: [
        { href: "/panel", label: "Yeni Sipariş", icon: "plus", short: "Yeni", exact: true, keywords: ["sipariş paneli", "sipariş oluştur", "form"] },
        { href: "/panel/siparisler", label: "Siparişler", icon: "list", short: "Siparişler", badge: "acik", keywords: ["sipariş listesi", "aktif", "takip"] },
        { href: "/panel/siparisler/tamamlanan", label: "Tamamlananlar", icon: "check-circle", keywords: ["arşiv", "tamamlanan siparişler"] },
      ],
    });
    groups.push({
      title: "Perakende",
      items: [
        { href: "/panel/perakende", label: "Online Çerçeve", icon: "frame", short: "Perakende", exact: true, keywords: ["çerçeveletme", "sihirbaz", "paspartu", "cam", "baskı"] },
        { href: "/panel/perakende/siparisler", label: "Perakende Siparişler", icon: "image", badge: "perakendeAcik", keywords: ["prk", "perakende sipariş"] },
        { href: "/panel/perakende/musteriler", label: "Perakende Müşteriler", icon: "user", keywords: ["perakende müşteri defteri"] },
      ],
    });
    groups.push({
      title: "Müşteri & İletişim",
      items: [
        { href: "/etiket", label: "Müşteriler & Etiket", icon: "users", keywords: ["müşteri", "kargo etiketi", "bayi", "cari"] },
        { href: "/panel/sms", label: "SMS Gönder", icon: "message", keywords: ["sms", "mesaj", "netgsm"] },
      ],
    });

    const yonetim: NavItem[] = [
      { href: "/panel/stok", label: "Stok Yükle", icon: "upload", keywords: ["excel", "günlük stok", "yükle"] },
    ];
    if (f.kur) yonetim.push({ href: "/panel/kur", label: "Günlük Kur", icon: "dollar", keywords: ["dolar", "euro", "kur"] });
    if (f.finance) {
      yonetim.push({
        href: "/panel/finans",
        label: "Finans",
        icon: "wallet",
        exact: true,
        keywords: ["kasa", "tahsilat", "gider"],
        children: [
          { href: "/panel/finans/kasa", label: "Kasa Raporu", icon: "credit-card", keywords: ["kasa"] },
          { href: "/panel/finans/giderler", label: "Giderler", icon: "arrow-down", keywords: ["gider", "masraf"] },
          { href: "/panel/finans/ceksenet", label: "Çek / Senet", icon: "file-text", keywords: ["çek", "senet", "vade"] },
          { href: "/panel/finans/personel", label: "Personel", icon: "briefcase", keywords: ["maaş", "personel"] },
        ],
      });
    }
    if (f.maliyet) yonetim.push({ href: "/panel/maliyet", label: "Maliyet & Kârlılık", icon: "percent", keywords: ["alış fiyatı", "kâr", "konteyner"] });
    if (f.raporlar) yonetim.push({ href: "/panel/raporlar", label: "Raporlar", icon: "bar-chart", keywords: ["ciro", "rapor", "analiz"] });
    groups.push({ title: "Yönetim", items: yonetim });
  }

  return groups;
}

/** Bütün menü öğeleri (alt öğeler dahil) düz liste. */
export function flatNav(f: NavFlags): NavItem[] {
  const out: NavItem[] = [];
  for (const g of navGroups(f)) {
    for (const it of g.items) {
      out.push(it);
      if (it.children) out.push(...it.children);
    }
  }
  return out;
}

/** Menüde olmayan detay sayfalarının başlıkları (kırıntı için). */
const EXTRA_TITLES: Record<string, string> = {
  "/panel/siparisler/detay": "Sipariş Detayı",
  "/panel/siparisler/duzenle": "Sipariş Düzenle",
  "/panel/perakende/siparisler/detay": "Perakende Sipariş Detayı",
  "/musteri": "Müşteri Cari Hesabı",
  "/giris": "Giriş",
};

export function isActive(item: NavItem, pathname: string): boolean {
  if (item.href === "/") return pathname === "/";
  if (item.exact) return pathname === item.href;
  return pathname === item.href || pathname.startsWith(item.href + "/");
}

/** Yolla en iyi eşleşen menü öğesinin href'i (en uzun ön ek kazanır). */
export function activeHref(f: NavFlags, pathname: string): string | null {
  let best: NavItem | null = null;
  for (const it of flatNav(f)) {
    if (!isActive(it, pathname)) continue;
    if (!best || it.href.length > best.href.length) best = it;
  }
  // "exact" bir öğe, kendi alt yollarını da sahiplensin (örn. /panel/siparisler/detay → Siparişler)
  if (!best) {
    for (const it of flatNav(f)) {
      if (it.href !== "/" && pathname.startsWith(it.href + "/")) {
        if (!best || it.href.length > best.href.length) best = it;
      }
    }
  }
  return best ? best.href : null;
}

export interface Crumb { label: string; href: string; }

// Bu yollar kendi başına sayfadır ama alt yolların "üstü" değildir:
// /panel = Yeni Sipariş formu, /portal = stok sorgusu. Kırıntıda ara halka olmazlar.
const NOT_A_PARENT = new Set(["/panel", "/portal"]);

export function breadcrumbsFor(f: NavFlags, pathname: string): Crumb[] {
  const titles = new Map<string, string>();
  for (const it of flatNav(f)) titles.set(it.href, it.label);
  for (const [k, v] of Object.entries(EXTRA_TITLES)) titles.set(k, v);

  const crumbs: Crumb[] = [{ label: "Ana Sayfa", href: "/" }];
  if (pathname === "/") return crumbs;

  const segs = pathname.split("/").filter(Boolean);
  let acc = "";
  for (let i = 0; i < segs.length; i++) {
    acc += "/" + segs[i];
    const t = titles.get(acc);
    if (t && (acc === pathname || !NOT_A_PARENT.has(acc))) {
      crumbs.push({ label: t, href: acc });
    } else if (acc.startsWith("/kataloglar/") && i === segs.length - 1) {
      crumbs.push({ label: "Katalog Görüntüleyici", href: acc });
    }
  }
  // Yol tanınmadıysa son parçayı başlık yap
  if (crumbs.length === 1) {
    const last = segs[segs.length - 1] || "";
    crumbs.push({ label: last.charAt(0).toUpperCase() + last.slice(1), href: pathname });
  }
  return crumbs;
}

/** Telefon alt sekme çubuğu: en çok kullanılan 4 hedef + Menü. */
export function tabItems(f: NavFlags): NavItem[] {
  if (f.role === "staff") {
    return [
      { href: "/", label: "Panel", icon: "home", exact: true },
      { href: "/panel/siparisler", label: "Siparişler", icon: "list", badge: "acik" },
      { href: "/panel", label: "Yeni", icon: "plus", exact: true },
      { href: "/panel/perakende", label: "Perakende", icon: "frame" },
    ];
  }
  return [
    { href: "/", label: "Ana Sayfa", icon: "home", exact: true },
    { href: "/portal", label: "Stok", icon: "package", exact: true },
    { href: "/portal/fiyat-listesi", label: "Fiyat", icon: "tag" },
    { href: "/kataloglar", label: "Katalog", icon: "book" },
  ];
}

export function initials(name: string): string {
  const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "?";
  const a = parts[0][0] || "";
  const b = parts.length > 1 ? parts[parts.length - 1][0] || "" : "";
  return (a + b).toLocaleUpperCase("tr-TR");
}
