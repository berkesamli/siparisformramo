// Gösterge paneli verisi — ucuz kaynaklardan (aylık indeksler, tek stok
// dosyası, günün kuru) derlenir; sipariş dosyaları tek tek OKUNMAZ.

import {
  readOrderIndex,
  rebuildOrderIndexFrom,
  listAllOrders,
  istanbulDateKey,
  lastNDateKeys,
  getDailyRates,
  blobConfigured,
  type OrderIndexEntry,
  type OrderStatus,
  type DailyRates,
} from "./orders";
import { readRetailIndexOrRebuild, type RetailIndexEntry } from "./retail-orders";
import { getStockData } from "./stock-store";
import { listCekSenet } from "./ceksenet";
import { memo } from "./server-cache";
import { FRAME_PROFILES } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import type { RetailStatus } from "@/data/perakende";

export interface DashFlags {
  finance: boolean;
  kur: boolean;
  owner: boolean;
}

export interface LiteStats {
  bugun: number;
  bugunToptan: number;
  bugunPerakende: number;
  dun: number;
  acik: number;
  perakendeAcik: number;
  kontrolsuz: number;
  ayAdet: number;
  ayCiro?: number;
  blob: boolean;
}

export interface SeriGun {
  date: string;   // YYYY-MM-DD
  label: string;  // "13 Eyl"
  gun: string;    // "Cmt"
  toptan: number;
  perakende: number;
  toptanCiro: number;
  perakendeCiro: number;
}

export interface Uyari {
  tip: "kontrol" | "stok" | "kur" | "cek" | "blob" | "acik";
  seviye: "info" | "warn" | "err";
  baslik: string;
  metin: string;
  href: string;
}

export interface StaffDashboard {
  role: "staff";
  blob: boolean;
  today: string;
  lite: LiteStats;
  kpi: {
    bugun: { toptan: number; perakende: number; toplam: number; dun: number };
    acik: { toptan: number; perakende: number };
    kontrolsuz: number;
    ay: { toptanAdet: number; perakendeAdet: number; toptanCiro: number; perakendeCiro: number };
    gun14: { toptan: number; perakende: number; toptanCiro: number; perakendeCiro: number };
  };
  seri14: SeriGun[];
  durum14: { toptan: Record<OrderStatus, number>; perakende: Record<RetailStatus, number> };
  sonToptan: OrderIndexEntry[];
  sonPerakende: RetailIndexEntry[];
  uyarilar: Uyari[];
  stok: { updatedAt: string; kalem: number; kaynak: string } | null;
  kur: DailyRates | null;
  cek?: { yaklasanAdet: number; yaklasanToplam: number; gecmisAdet: number };
  hesaplandi: string;
}

export interface CustomerDashboard {
  role: "customer";
  today: string;
  stok: { updatedAt: string; kalem: number } | null;
  katalog: { profil: number; teknik: number; seri: { seri: string; adet: number }[] };
  hesaplandi: string;
}

const AY_KISA = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"];
const GUN_KISA = ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"];
const r2 = (n: number) => Math.round(n * 100) / 100;

function gunEtiketi(dateKey: string): { label: string; gun: string } {
  const [y, m, d] = dateKey.split("-").map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  return { label: `${d} ${AY_KISA[m - 1]}`, gun: GUN_KISA[dt.getUTCDay()] };
}

const TOPTAN_ACIK: OrderStatus[] = ["olusturuldu", "hazirlaniyor"];
const PERAKENDE_ACIK: RetailStatus[] = ["Beklemede", "Hazırlanıyor", "Hazır"];

async function toptanIndeks(): Promise<OrderIndexEntry[]> {
  return memo("dash:toptan-idx", 45_000, async () => {
    const idx = await readOrderIndex(2);
    if (idx.length || !blobConfigured()) return idx;
    // İlk kurulum: indeks yoksa tam taramadan kur (bir defalık)
    const hepsi = await listAllOrders();
    if (!hepsi.length) return [];
    await rebuildOrderIndexFrom(hepsi);
    return readOrderIndex(2);
  });
}

async function perakendeIndeks(): Promise<RetailIndexEntry[]> {
  return memo("dash:perakende-idx", 45_000, () => readRetailIndexOrRebuild(2));
}

export async function computeStaffDashboard(flags: DashFlags): Promise<StaffDashboard> {
  const today = istanbulDateKey();
  const keys14 = lastNDateKeys(14); // bugün → 13 gün önce
  const set14 = new Set(keys14);
  const set7 = new Set(lastNDateKeys(7));
  const dun = keys14[1];
  const ay = today.slice(0, 7);
  const blob = blobConfigured();

  const [tIdx, pIdx, stokData, kur, cekler] = await Promise.all([
    toptanIndeks(),
    perakendeIndeks(),
    memo("dash:stok", 60_000, () => getStockData()).catch(() => null),
    memo(`dash:kur:${today}`, 60_000, () => getDailyRates(today)).catch(() => null),
    flags.finance ? memo("dash:cek", 60_000, () => listCekSenet()).catch(() => []) : Promise.resolve([]),
  ]);

  const tAktif = tIdx.filter((o) => o.status !== "iptal");
  const pAktif = pIdx.filter((o) => o.status !== "İptal");

  const bugunToptan = tAktif.filter((o) => o.dateKey === today).length;
  const bugunPerakende = pAktif.filter((o) => o.dateKey === today).length;
  const dunToplam = tAktif.filter((o) => o.dateKey === dun).length + pAktif.filter((o) => o.dateKey === dun).length;
  const acikToptan = tIdx.filter((o) => TOPTAN_ACIK.includes(o.status)).length;
  const acikPerakende = pIdx.filter((o) => PERAKENDE_ACIK.includes(o.status)).length;
  const kontrolsuz = tIdx.filter((o) => set7.has(o.dateKey) && o.status !== "iptal" && o.kontrol === false).length;

  const ayT = tAktif.filter((o) => o.dateKey.startsWith(ay));
  const ayP = pAktif.filter((o) => o.dateKey.startsWith(ay));
  const ayToptanCiro = r2(ayT.reduce((s, o) => s + (Number(o.net) || 0), 0));
  const ayPerakendeCiro = r2(ayP.reduce((s, o) => s + (Number(o.total) || 0), 0));

  // 14 günlük seri (eskiden yeniye)
  const seri14: SeriGun[] = [...keys14].reverse().map((k) => {
    const t = tAktif.filter((o) => o.dateKey === k);
    const p = pAktif.filter((o) => o.dateKey === k);
    const { label, gun } = gunEtiketi(k);
    return {
      date: k,
      label,
      gun,
      toptan: t.length,
      perakende: p.length,
      toptanCiro: r2(t.reduce((s, o) => s + (Number(o.net) || 0), 0)),
      perakendeCiro: r2(p.reduce((s, o) => s + (Number(o.total) || 0), 0)),
    };
  });
  const gun14 = seri14.reduce(
    (a, g) => ({
      toptan: a.toptan + g.toptan,
      perakende: a.perakende + g.perakende,
      toptanCiro: r2(a.toptanCiro + g.toptanCiro),
      perakendeCiro: r2(a.perakendeCiro + g.perakendeCiro),
    }),
    { toptan: 0, perakende: 0, toptanCiro: 0, perakendeCiro: 0 }
  );

  const durumT: Record<OrderStatus, number> = { olusturuldu: 0, hazirlaniyor: 0, tamamlandi: 0, iptal: 0 };
  for (const o of tIdx) if (set14.has(o.dateKey)) durumT[o.status] = (durumT[o.status] || 0) + 1;
  const durumP: Record<RetailStatus, number> = { Beklemede: 0, "Hazırlanıyor": 0, "Hazır": 0, "Teslim Edildi": 0, "İptal": 0 };
  for (const o of pIdx) if (set14.has(o.dateKey)) durumP[o.status] = (durumP[o.status] || 0) + 1;

  // Uyarılar
  const uyarilar: Uyari[] = [];
  if (!blob) {
    uyarilar.push({
      tip: "blob",
      seviye: "info",
      baslik: "Kalıcı depolama bağlı değil",
      metin: "Vercel Blob yapılandırılana kadar sipariş ve müşteri kayıtları saklanmaz; sayaçlar 0 görünür.",
      href: "/panel/stok",
    });
  }
  if (kontrolsuz > 0) {
    uyarilar.push({
      tip: "kontrol",
      seviye: "warn",
      baslik: `${kontrolsuz} sipariş kontrol bekliyor`,
      metin: "Son 7 günde merkez kontrolü yapılmamış toptan siparişler var.",
      href: "/panel/siparisler",
    });
  }
  if (flags.kur && blob && !kur) {
    uyarilar.push({
      tip: "kur",
      seviye: "warn",
      baslik: "Bugünün kuru girilmedi",
      metin: "Dolar / euro kuru belirlenmeden alınan siparişler ilk girilen kuru kullanır.",
      href: "/panel/kur",
    });
  }
  let stok: StaffDashboard["stok"] = null;
  if (stokData) {
    stok = { updatedAt: stokData.updatedAt, kalem: stokData.items.length, kaynak: stokData.sourceName };
    const yasSaat = (Date.now() - new Date(stokData.updatedAt).getTime()) / 3_600_000;
    if (yasSaat > 36) {
      const gun = Math.floor(yasSaat / 24);
      uyarilar.push({
        tip: "stok",
        seviye: "warn",
        baslik: "Stok verisi güncel değil",
        metin: `Son stok yüklemesi ${gun} gün önce yapıldı. Günlük Excel'i yükleyin.`,
        href: "/panel/stok",
      });
    }
  }
  let cek: StaffDashboard["cek"];
  if (flags.finance) {
    const portfoy = (cekler as { durum: string; vade: string; tutar: number }[]).filter((c) => c.durum === "portfoyde");
    const otuz = istanbulDateKey(new Date(Date.now() + 30 * 86_400_000));
    const yaklasan = portfoy.filter((c) => c.vade <= otuz);
    cek = {
      yaklasanAdet: yaklasan.length,
      yaklasanToplam: r2(yaklasan.reduce((s, c) => s + (Number(c.tutar) || 0), 0)),
      gecmisAdet: portfoy.filter((c) => c.vade < today).length,
    };
    if (cek.gecmisAdet > 0) {
      uyarilar.push({
        tip: "cek",
        seviye: "err",
        baslik: `${cek.gecmisAdet} çek/senet vadesi geçti`,
        metin: "Portföyde vadesi geçmiş ama işlem görmemiş evrak var.",
        href: "/panel/finans/ceksenet",
      });
    } else if (cek.yaklasanAdet > 0) {
      uyarilar.push({
        tip: "cek",
        seviye: "info",
        baslik: `${cek.yaklasanAdet} çek/senet 30 gün içinde vadeli`,
        metin: `Toplam ₺${cek.yaklasanToplam.toLocaleString("tr-TR")} — takvimi finans ekranından izleyin.`,
        href: "/panel/finans/ceksenet",
      });
    }
  }
  if (acikToptan + acikPerakende > 25) {
    uyarilar.push({
      tip: "acik",
      seviye: "info",
      baslik: `${acikToptan + acikPerakende} açık sipariş`,
      metin: "Tamamlanan siparişlerin durumunu güncelleyin; liste daha sade kalır.",
      href: "/panel/siparisler",
    });
  }

  const lite: LiteStats = {
    bugun: bugunToptan + bugunPerakende,
    bugunToptan,
    bugunPerakende,
    dun: dunToplam,
    acik: acikToptan,
    perakendeAcik: acikPerakende,
    kontrolsuz,
    ayAdet: ayT.length + ayP.length,
    ayCiro: flags.finance ? r2(ayToptanCiro + ayPerakendeCiro) : undefined,
    blob,
  };

  return {
    role: "staff",
    blob,
    today,
    lite,
    kpi: {
      bugun: { toptan: bugunToptan, perakende: bugunPerakende, toplam: bugunToptan + bugunPerakende, dun: dunToplam },
      acik: { toptan: acikToptan, perakende: acikPerakende },
      kontrolsuz,
      ay: { toptanAdet: ayT.length, perakendeAdet: ayP.length, toptanCiro: ayToptanCiro, perakendeCiro: ayPerakendeCiro },
      gun14,
    },
    seri14,
    durum14: { toptan: durumT, perakende: durumP },
    sonToptan: tIdx.slice(0, 8),
    sonPerakende: pIdx.slice(0, 6),
    uyarilar,
    stok,
    kur,
    cek,
    hesaplandi: new Date().toISOString(),
  };
}

export async function computeCustomerDashboard(): Promise<CustomerDashboard> {
  const today = istanbulDateKey();
  const stokData = await memo("dash:stok", 60_000, () => getStockData()).catch(() => null);
  const seriMap = new Map<string, number>();
  for (const p of FRAME_PROFILES) seriMap.set(p.series, (seriMap.get(p.series) || 0) + 1);
  return {
    role: "customer",
    today,
    stok: stokData ? { updatedAt: stokData.updatedAt, kalem: stokData.items.length } : null,
    katalog: {
      profil: FRAME_PROFILES.length,
      teknik: TECHNICAL_PRODUCTS.length,
      seri: [...seriMap.entries()].map(([seri, adet]) => ({ seri, adet })),
    },
    hesaplandi: new Date().toISOString(),
  };
}
