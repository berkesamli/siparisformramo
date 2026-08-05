// Alış fiyatları ve maliyet — PARTİ (konteyner) bazlı. finans/maliyetler.json.
//
// Her gelen konteynerin maliyeti farklıdır: parti açılır (ad + geliş tarihi),
// o partinin birim alış fiyatları girilir; partinin genel gider yüzdesi
// (nakliye, gümrük, fire) sonradan da girilebilir — yüzde girilmeden kâr
// hesaplanmaz, "yüzde bekliyor" görünür.
//
// Analizde bir satışın maliyeti, sipariş tarihinden ÖNCEKİ EN SON partide o
// kod için girilen fiyattan hesaplanır (satış hangi partinin malından
// yapıldıysa ona en yakın doğru varsayım). Sipariş ilk partiden eskiyse ilk
// partinin fiyatı kullanılır.
//
// Erişim: yalnızca MALIYET_USERNAMES (varsayılan: berke, özgür).

import { blobConfigured } from "./orders";

export type AlisBirimi = "USD" | "EUR" | "TL";

export interface PartiKalemi {
  code: string; // görünen kod ("4501 S")
  alis: number; // birim alış (/mt)
  currency: AlisBirimi;
}

export interface Parti {
  id: string; // PT-xxxxx
  ad: string; // örn. "Ağustos 2026 — 40'lık"
  tarih: string; // geliş tarihi YYYY-MM-DD
  /** Genel gider yüzdesi — null: henüz girilmedi, kâr hesaplanmaz. */
  pct: number | null;
  note?: string;
  createdAt: string;
  by: string;
  items: Record<string, PartiKalemi>; // anahtar: normKod
}

export interface MaliyetData {
  updatedAt: string;
  partiler: Parti[];
}

const PATH = "finans/maliyetler.json";

export const normKod = (s: string) =>
  String(s || "").toUpperCase().replace(/\s+/g, "");

export function newPartiId(): string {
  return "PT" + Math.random().toString(36).slice(2, 8).toUpperCase();
}

export async function getMaliyetData(): Promise<MaliyetData> {
  const bos: MaliyetData = { updatedAt: "", partiler: [] };
  if (!blobConfigured()) return bos;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(PATH, { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return bos;
    const d = JSON.parse(await new Response(r.stream).text()) as MaliyetData;
    return d.partiler ? d : bos; // eski şema gelirse boş başla
  } catch {
    return bos;
  }
}

export async function saveMaliyetData(d: MaliyetData): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(PATH, JSON.stringify(d), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export interface MaliyetSecim {
  parti: Parti;
  kalem: PartiKalemi;
}

/**
 * Kod → tarih sıralı kalem listesi çıkarır; sipariş tarihine göre geçerli
 * partiyi seçmekte kullanılır.
 */
export function kodIndeksi(d: MaliyetData): Map<string, MaliyetSecim[]> {
  const map = new Map<string, MaliyetSecim[]>();
  const sirali = [...d.partiler].sort((a, b) => a.tarih.localeCompare(b.tarih));
  for (const parti of sirali) {
    for (const [nk, kalem] of Object.entries(parti.items)) {
      const list = map.get(nk) || [];
      list.push({ parti, kalem });
      map.set(nk, list);
    }
  }
  return map;
}

/** Sipariş tarihi için geçerli kalem: tarih ≤ sipariş olan son parti, yoksa ilk. */
export function gecerliKalem(
  list: MaliyetSecim[] | undefined,
  orderDateKey: string
): MaliyetSecim | null {
  if (!list || !list.length) return null;
  let secim: MaliyetSecim | null = null;
  for (const s of list) {
    if (s.parti.tarih <= orderDateKey) secim = s;
    else break;
  }
  return secim || list[0];
}

/** Birim maliyet TL — partinin yüzdesi girilmemişse null. */
export function birimMaliyetTL(
  s: MaliyetSecim,
  usdRate: number,
  eurRate: number
): number | null {
  if (s.parti.pct == null) return null;
  const k = s.kalem;
  const kur = k.currency === "USD" ? usdRate : k.currency === "EUR" ? eurRate : 1;
  return k.alis * kur * (1 + s.parti.pct / 100);
}
