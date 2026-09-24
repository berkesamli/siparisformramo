// Çalışanın KENDİ satışları — yalnızca oturumdaki kullanıcının adıyla girilmiş
// siparişler (toptan `net`, perakende `total`; iptaller hariç — raporlarla aynı
// tanım). Aylık indekslerden hesaplanır; sipariş dosyaları tek tek okunmaz.
// Başka çalışanların kişisel rakamları bu modülden dönmez.
//
// Bölge satışları: kullanıcı bir bölgenin sorumlusuysa (BOLGE_SORUMLULARI /
// lib/customers bolgeler()), o bölgedeki müşterilerin TÜM siparişleri ayrıca
// "bölge satışları" olarak döner — siparişi kim almış olursa olsun; alan
// çalışanın adı sipariş satırında görünür.

import {
  readOrderIndex,
  rebuildOrderIndexFrom,
  listAllOrders,
  istanbulDateKey,
  lastNDateKeys,
  blobConfigured,
  STATUS_LABELS,
  type OrderIndexEntry,
} from "./orders";
import { readRetailIndexOrRebuild, type RetailIndexEntry } from "./retail-orders";
import { normalizeUsername } from "@/data/users";
import { memo } from "./server-cache";
import type { SeriGun } from "./dashboard";
import { bolgeCozucu, kullanicininBolgesi } from "./bolge-atama";
import { bolgeler, type Bolge } from "./customers";

const AY_KISA = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"];
const AY_UZUN = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];
const GUN_KISA = ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"];
const r2 = (n: number) => Math.round(n * 100) / 100;
const AY_SAYISI = 6; // geriye dönük kaç ay okunur

export interface DonemOzet {
  adet: number;
  ciro: number;
  toptanAdet: number;
  perakendeAdet: number;
  toptanCiro: number;
  perakendeCiro: number;
}

export interface AyOzet extends DonemOzet {
  ay: string;    // YYYY-MM
  label: string; // "Eylül 2026"
  kisa: string;  // "Eyl"
}

export interface KisiselSiparis {
  tur: "toptan" | "perakende";
  orderId: string;
  dateKey: string;
  createdAt: string;
  musteri: string;
  tutar: number;
  status: string;      // Türkçe etiket
  statusKind: "warn" | "info" | "ok" | "err" | "yarim" | "";
  href: string;
  alan?: string;       // siparişi alan çalışan (bölge listesinde)
}

/** Bölge sorumlusunun bölgesindeki tüm satışlar. */
export interface BolgeSatis {
  id: Bolge;
  label: string;
  bugun: DonemOzet;
  hafta: DonemOzet;
  secili: DonemOzet;
  oncekiAy: DonemOzet;
  benimCiro: number;      // seçili ayda bu çalışanın kendi aldığı bölge siparişleri
  baskalariCiro: number;  // seçili ayda başkalarının aldığı
  aylar: AyOzet[];
  sonSiparisler: KisiselSiparis[];
}

export interface KisiselOzet {
  employee: string;
  today: string;
  ay: string;                 // seçili ay (YYYY-MM)
  ayLabel: string;
  aylar: AyOzet[];            // son 6 ay, eskiden yeniye (seçilebilir aralık)
  bugun: DonemOzet;
  hafta: DonemOzet;           // son 7 gün
  secili: DonemOzet;          // seçili ay
  oncekiAy: DonemOzet;        // seçili ayın bir öncesi (kıyas)
  ortalamaSiparis: number;    // seçili ayda sipariş başına ciro
  enIyiGun: { date: string; label: string; ciro: number } | null; // seçili ay
  seri14: SeriGun[];          // son 14 gün, yalnızca bu çalışan
  sonSiparisler: KisiselSiparis[]; // seçili aydaki son 30 sipariş
  bolge?: BolgeSatis | null;       // bölge sorumlusuysa bölgesinin satışları
  blob: boolean;
  hesaplandi: string;
}

const bosDonem = (): DonemOzet => ({ adet: 0, ciro: 0, toptanAdet: 0, perakendeAdet: 0, toptanCiro: 0, perakendeCiro: 0 });

function topla(t: OrderIndexEntry[], p: RetailIndexEntry[]): DonemOzet {
  const d = bosDonem();
  for (const o of t) { d.toptanAdet++; d.toptanCiro += Number(o.net) || 0; }
  for (const o of p) { d.perakendeAdet++; d.perakendeCiro += Number(o.total) || 0; }
  d.toptanCiro = r2(d.toptanCiro);
  d.perakendeCiro = r2(d.perakendeCiro);
  d.adet = d.toptanAdet + d.perakendeAdet;
  d.ciro = r2(d.toptanCiro + d.perakendeCiro);
  return d;
}

function ayKeyleri(today: string, n: number): string[] {
  const [y, m] = today.split("-").map(Number);
  const out: string[] = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(Date.UTC(y, m - 1 - i, 1));
    out.push(`${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`);
  }
  return out;
}

function ayEtiket(ay: string): { label: string; kisa: string } {
  const [y, m] = ay.split("-").map(Number);
  return { label: `${AY_UZUN[m - 1]} ${y}`, kisa: AY_KISA[m - 1] };
}

function gunEtiketi(dateKey: string): { label: string; gun: string } {
  const [y, m, d] = dateKey.split("-").map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  return { label: `${d} ${AY_KISA[m - 1]}`, gun: GUN_KISA[dt.getUTCDay()] };
}

const TOPTAN_KIND: Record<string, KisiselSiparis["statusKind"]> = { olusturuldu: "warn", hazirlaniyor: "info", yarim: "yarim", tamamlandi: "ok", iptal: "err" };
const PERAKENDE_KIND: Record<string, KisiselSiparis["statusKind"]> = { Beklemede: "warn", "Hazırlanıyor": "info", "Hazır": "ok", "Teslim Edildi": "", "İptal": "err" };

/** Son 6 ayın indeksleri — kullanıcıdan bağımsız, kısa süre paylaşımlı önbellek. */
async function indeksler(): Promise<{ t: OrderIndexEntry[]; p: RetailIndexEntry[] }> {
  const t = await memo("cirom:toptan-idx", 45_000, async () => {
    const idx = await readOrderIndex(AY_SAYISI);
    if (idx.length || !blobConfigured()) return idx;
    const hepsi = await listAllOrders();
    if (!hepsi.length) return [];
    await rebuildOrderIndexFrom(hepsi);
    return readOrderIndex(AY_SAYISI);
  });
  const p = await memo("cirom:perakende-idx", 45_000, () => readRetailIndexOrRebuild(AY_SAYISI));
  return { t, p };
}

export async function computePersonalSales(employeeName: string, aySecim?: string, username?: string): Promise<KisiselOzet> {
  const today = istanbulDateKey();
  const aylar = ayKeyleri(today, AY_SAYISI);
  const ay = aySecim && aylar.includes(aySecim) ? aySecim : today.slice(0, 7);
  const oncekiAyKey = ayKeyleri(ay + "-01", 2)[0];
  const ben = normalizeUsername(employeeName);
  const benim = (e: string | undefined) => normalizeUsername(e || "") === ben;

  const { t: tHepsi, p: pHepsi } = await indeksler();
  // Yalnızca bu çalışanın, iptal olmayan siparişleri
  const t = tHepsi.filter((o) => benim(o.employee) && o.status !== "iptal");
  const p = pHepsi.filter((o) => benim(o.employee) && o.status !== "İptal");

  const keys14 = lastNDateKeys(14);
  const set7 = new Set(keys14.slice(0, 7));
  const bugun = topla(t.filter((o) => o.dateKey === today), p.filter((o) => o.dateKey === today));
  const hafta = topla(t.filter((o) => set7.has(o.dateKey)), p.filter((o) => set7.has(o.dateKey)));
  const tAy = t.filter((o) => o.dateKey.startsWith(ay));
  const pAy = p.filter((o) => o.dateKey.startsWith(ay));
  const secili = topla(tAy, pAy);
  const oncekiAy = topla(t.filter((o) => o.dateKey.startsWith(oncekiAyKey)), p.filter((o) => o.dateKey.startsWith(oncekiAyKey)));

  const aylarOzet: AyOzet[] = aylar.map((a) => ({
    ay: a,
    ...ayEtiket(a),
    ...topla(t.filter((o) => o.dateKey.startsWith(a)), p.filter((o) => o.dateKey.startsWith(a))),
  }));

  // Seçili ayın en iyi günü
  const gunMap = new Map<string, number>();
  for (const o of tAy) gunMap.set(o.dateKey, (gunMap.get(o.dateKey) || 0) + (Number(o.net) || 0));
  for (const o of pAy) gunMap.set(o.dateKey, (gunMap.get(o.dateKey) || 0) + (Number(o.total) || 0));
  let enIyiGun: KisiselOzet["enIyiGun"] = null;
  for (const [d, c] of gunMap) if (!enIyiGun || c > enIyiGun.ciro) enIyiGun = { date: d, label: gunEtiketi(d).label, ciro: r2(c) };

  const seri14: SeriGun[] = [...keys14].reverse().map((k) => {
    const tt = t.filter((o) => o.dateKey === k);
    const pp = p.filter((o) => o.dateKey === k);
    const { label, gun } = gunEtiketi(k);
    return {
      date: k, label, gun,
      toptan: tt.length,
      perakende: pp.length,
      toptanCiro: r2(tt.reduce((s, o) => s + (Number(o.net) || 0), 0)),
      perakendeCiro: r2(pp.reduce((s, o) => s + (Number(o.total) || 0), 0)),
    };
  });

  const sonSiparisler = siparisListesi(tAy, pAy);

  // Bölge satışları (sorumluysa): bölgedeki müşterilerin tüm siparişleri
  let bolge: BolgeSatis | null = null;
  const bolgeId = kullanicininBolgesi(employeeName, username);
  if (bolgeId) {
    const coz = await bolgeCozucu().catch(() => null);
    if (coz) {
      const tB = tHepsi.filter((o) => o.status !== "iptal" && coz(o) === bolgeId);
      const pB = pHepsi.filter((o) => o.status !== "İptal" && coz({ customerId: o.customerId, customerName: o.customerName }) === bolgeId);
      const tBAy = tB.filter((o) => o.dateKey.startsWith(ay));
      const pBAy = pB.filter((o) => o.dateKey.startsWith(ay));
      const seciliB = topla(tBAy, pBAy);
      const benimCiro = topla(tBAy.filter((o) => benim(o.employee)), pBAy.filter((o) => benim(o.employee))).ciro;
      bolge = {
        id: bolgeId,
        label: bolgeler()[bolgeId].label,
        bugun: topla(tB.filter((o) => o.dateKey === today), pB.filter((o) => o.dateKey === today)),
        hafta: topla(tB.filter((o) => set7.has(o.dateKey)), pB.filter((o) => set7.has(o.dateKey))),
        secili: seciliB,
        oncekiAy: topla(tB.filter((o) => o.dateKey.startsWith(oncekiAyKey)), pB.filter((o) => o.dateKey.startsWith(oncekiAyKey))),
        benimCiro,
        baskalariCiro: r2(seciliB.ciro - benimCiro),
        aylar: aylar.map((a) => ({
          ay: a,
          ...ayEtiket(a),
          ...topla(tB.filter((o) => o.dateKey.startsWith(a)), pB.filter((o) => o.dateKey.startsWith(a))),
        })),
        sonSiparisler: siparisListesi(tBAy, pBAy),
      };
    }
  }

  return {
    employee: employeeName,
    today,
    ay,
    ayLabel: ayEtiket(ay).label,
    aylar: aylarOzet,
    bugun,
    hafta,
    secili,
    oncekiAy,
    ortalamaSiparis: secili.adet ? r2(secili.ciro / secili.adet) : 0,
    enIyiGun,
    seri14,
    sonSiparisler,
    bolge,
    blob: blobConfigured(),
    hesaplandi: new Date().toISOString(),
  };
}

/** Seçili ayın siparişleri, yeniden eskiye, en fazla 30 (alan çalışan adıyla). */
function siparisListesi(tAy: OrderIndexEntry[], pAy: RetailIndexEntry[]): KisiselSiparis[] {
  return [
    ...tAy.map((o): KisiselSiparis => ({
      tur: "toptan", orderId: o.orderId, dateKey: o.dateKey, createdAt: o.createdAt,
      musteri: o.customer || "—", tutar: Number(o.net) || 0,
      status: STATUS_LABELS[o.status], statusKind: TOPTAN_KIND[o.status] || "",
      href: `/panel/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
      alan: o.employee || undefined,
    })),
    ...pAy.map((o): KisiselSiparis => ({
      tur: "perakende", orderId: o.orderId, dateKey: o.dateKey, createdAt: o.createdAt,
      musteri: o.customerName || "—", tutar: Number(o.total) || 0,
      status: o.status, statusKind: PERAKENDE_KIND[o.status] || "",
      href: `/panel/perakende/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
      alan: o.employee || undefined,
    })),
  ]
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt))
    .slice(0, 30);
}
