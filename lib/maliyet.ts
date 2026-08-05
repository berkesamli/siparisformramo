// Alış fiyatları ve maliyet — finans/maliyetler.json (tek dosya).
//
// Yalnızca firma sahipleri görür ve düzenler. Maliyet hesabı yüzdeseldir:
//   birim maliyet = alış fiyatı × (1 + genel gider yüzdesi / 100)
// Genel gider yüzdesi (nakliye, gümrük, fire, işçilik payı) tek yerden
// tanımlanır; gerekirse kod bazında özel yüzde ile ezilebilir.
//
// Tek dosya tercihi bilinçli: ~200-400 kayıt küçük bir JSON'dur, okuma tek
// get'tir ve düzenleme yalnız 3 kişide olduğundan yarış riski yoktur.

import { blobConfigured } from "./orders";

export type AlisBirimi = "USD" | "EUR" | "TL";

export interface MaliyetKaydi {
  code: string; // profil/ürün kodu (renk eki OLMADAN: "4501 S", "GB211")
  alis: number; // birim alış fiyatı (profillerde /mt)
  currency: AlisBirimi;
  /** Kod bazlı özel genel gider yüzdesi — boşsa genel yüzde uygulanır. */
  pct?: number;
  note?: string;
  updatedAt: string;
  by: string;
}

export interface MaliyetData {
  updatedAt: string;
  /** Genel gider yüzdesi (%). Örn. 18 → maliyet = alış × 1,18 */
  defaultPct: number;
  items: Record<string, MaliyetKaydi>; // anahtar: normalize kod
}

const PATH = "finans/maliyetler.json";

export const normKod = (s: string) =>
  String(s || "").toUpperCase().replace(/\s+/g, "");

export async function getMaliyetData(): Promise<MaliyetData> {
  const bos: MaliyetData = { updatedAt: "", defaultPct: 0, items: {} };
  if (!blobConfigured()) return bos;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(PATH, { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return bos;
    return JSON.parse(await new Response(r.stream).text()) as MaliyetData;
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

/**
 * Birim maliyeti TL olarak hesaplar. Kur, siparişin kendi kuru olmalı ki
 * satışla maliyet aynı günün parasıyla karşılaştırılsın.
 */
export function birimMaliyetTL(
  k: MaliyetKaydi,
  defaultPct: number,
  usdRate: number,
  eurRate: number
): number {
  const kur = k.currency === "USD" ? usdRate : k.currency === "EUR" ? eurRate : 1;
  const pct = k.pct != null && k.pct >= 0 ? k.pct : defaultPct;
  return k.alis * kur * (1 + pct / 100);
}
