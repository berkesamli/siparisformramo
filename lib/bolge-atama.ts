// Siparişleri müşteri bölgesine (Ankara / İstanbul / Taşra) bağlar.
// Bölge müşteri kartından gelir (musteriBolgesi): siparişte customerId varsa
// doğrudan, yoksa müşteri adı kayıtlı bir kartla birebir eşleşiyorsa oradan.
// Hiçbiri tutmazsa "kayitsiz": müşteri deftere girilmemiş demektir; gösterge
// panelinde ayrı gösterilir ki eksik kayıtlar fark edilsin.
//
// Kullanım: gösterge paneli (bölge cirosu) ve Satışlarım (bölge sorumlusunun
// bölgesindeki tüm satışlar — siparişi kim almış olursa olsun).

import { listCustomers, musteriBolgesi, normalizeCity, customerTitle, bolgeler, type Bolge } from "./customers";
import { normalizeUsername } from "@/data/users";
import { memo } from "./server-cache";

export type BolgeKovasi = Bolge | "kayitsiz";
export const BOLGE_KOVALARI: BolgeKovasi[] = ["ankara", "istanbul", "tasra", "kayitsiz"];

export function kovaEtiketi(k: BolgeKovasi): string {
  if (k === "kayitsiz") return "Kayıtsız müşteri";
  return bolgeler()[k].label;
}

/** Ad karşılaştırma anahtarı: Türkçe karakter/büyük-küçük/boşluk duyarsız. */
const anahtar = (s: string) => normalizeCity(s).replace(/[^a-z0-9]/g, "");

export interface SiparisBolgeGirdisi {
  customerId?: string;
  customer?: string;      // toptan sipariş müşteri adı
  customerName?: string;  // perakende sipariş müşteri adı
}

export interface BolgeCozucu {
  (o: SiparisBolgeGirdisi): BolgeKovasi;
  musteriSayisi: number;
}

/** Müşteri defterinden id → bölge ve ad-anahtarı → bölge haritaları (60 sn önbellek). */
export async function bolgeCozucu(): Promise<BolgeCozucu> {
  return memo("bolge:cozucu", 60_000, async () => {
    const musteriler = await listCustomers().catch(() => []);
    const idMap = new Map<string, Bolge>();
    const adMap = new Map<string, Bolge>();
    for (const c of musteriler) {
      const b = musteriBolgesi(c);
      idMap.set(c.id, b);
      for (const ad of [customerTitle(c), c.company, `${c.firstName || ""} ${c.lastName || ""}`]) {
        const k = anahtar(ad || "");
        // Aynı ad iki kartta farklı bölgedeyse ilk kayıt kazanır; kısa anahtarlar (≤3) atlanır
        if (k.length > 3 && !adMap.has(k)) adMap.set(k, b);
      }
    }
    const coz = ((o: SiparisBolgeGirdisi): BolgeKovasi => {
      if (o.customerId && idMap.has(o.customerId)) return idMap.get(o.customerId)!;
      const k = anahtar(o.customer || o.customerName || "");
      return (k && adMap.get(k)) || "kayitsiz";
    }) as BolgeCozucu;
    coz.musteriSayisi = musteriler.length;
    return coz;
  });
}

/**
 * Kullanıcının sorumlu olduğu bölge (varsa). BOLGE_SORUMLULARI değeri
 * çalışanın adı ("Ramazan Kaypan") ya da kullanıcı adı ("ramazan") olabilir.
 */
export function kullanicininBolgesi(name: string, username?: string): Bolge | null {
  const adaylar = [name, username || ""].map((x) => normalizeUsername(x || "")).filter(Boolean);
  const tanim = bolgeler();
  for (const b of ["ankara", "istanbul", "tasra"] as Bolge[]) {
    const s = normalizeUsername(tanim[b].sorumlu || "");
    if (s && adaylar.includes(s)) return b;
  }
  return null;
}
