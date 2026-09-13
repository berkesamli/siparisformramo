// Katalog yardımcıları — VERİ İÇERMEZ.
// Toptan fiyat listesi (data/catalog, data/technical) istemci paketine
// girmesin diye tipler ve saf fonksiyonlar buraya ayrıldı: fonksiyonlar
// listeyi parametre alır. Sunucu tarafı data/catalog'daki listeyle, istemci
// tarafı /api/katalog'dan (oturumla) çektiği listeyle çağırır.

export type StockStatus = "var" | "az" | "yok";

export interface FrameProfile {
  code: string;
  series: string;
  koliAdet: number;
  koliMetraj: number; // MT
  priceUSD: number; // USD/mt toptan liste fiyatı
  stok: StockStatus;
}

export interface TechnicalProduct {
  code: string;
  name: string;
  category: string;
  adetPerKutu: number;
  priceEUR?: number;
  priceTL?: number;
  isKarton?: boolean;
  stok?: StockStatus;
  // Ürün görseli (tam URL). Sipariş formundaki seçicide küçük önizleme
  // olarak çıkar. Boş bırakılabilir — görseller zamanla eklenir.
  image?: string;
}

export const SERIES_ORDER = ["A", "G", "KS", "N", "T", "W", "Y"];

export function boyLength(profile: FrameProfile): number {
  return profile.koliAdet > 0 ? profile.koliMetraj / profile.koliAdet : 0;
}

/**
 * Metrajı "X koli + Y boy" metnine çevirir (fiş/PDF gösterimi için).
 * Örn. 145 mt, koli 72,5 mt → "2 koli"; 160,1 mt → "2 koli + 5 boy".
 */
export function koliBoyText(metres: number, profile: FrameProfile): string {
  const koliM = profile.koliMetraj;
  const boyM = boyLength(profile) || 2.9;
  if (metres <= 0 || koliM <= 0 || boyM <= 0) return "";
  const EPS = 0.01;
  let koli = Math.floor((metres + EPS) / koliM);
  const remainder = metres - koli * koliM;
  let boy = Math.round(remainder / boyM);
  // Kalan boylar tam bir koliyi tamamlıyorsa yukarı yuvarla
  if (boy >= profile.koliAdet) {
    koli += 1;
    boy = 0;
  }
  const parts: string[] = [];
  if (koli > 0) parts.push(`${koli} koli`);
  if (boy > 0) parts.push(`${boy} boy`);
  if (!parts.length) return "";
  return parts.join(" + ");
}

/** Kodu listede arar; "GC065-1473BX" gibi renkli kodlarda taban koda düşer. */
export function profilBul(
  list: FrameProfile[],
  code: string
): FrameProfile | undefined {
  const norm = (s: string) => s.toUpperCase().replace(/\s+/g, "");
  const q = norm(code);
  if (!q) return undefined;
  const exact = list.find((f) => norm(f.code) === q);
  if (exact) return exact;
  const base = q.split("-")[0];
  if (base && base !== q) {
    return list.find((f) => norm(f.code) === base);
  }
  return undefined;
}

export function teknikBul(
  list: TechnicalProduct[],
  code: string
): TechnicalProduct | undefined {
  return list.find((t) => t.code === code);
}

export function teknikKategorili(
  list: TechnicalProduct[]
): Record<string, TechnicalProduct[]> {
  const map: Record<string, TechnicalProduct[]> = {};
  for (const t of list) {
    (map[t.category] ||= []).push(t);
  }
  return map;
}
