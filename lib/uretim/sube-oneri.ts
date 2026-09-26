// Stoktan şube önerisi — işin kalemlerindeki profil kodlarının Ankara ve
// İstanbul deposundaki boyuna bakar, müşterinin şehrini de gözeterek işin
// hangi mağazada yapılacağını önerir. Öneri bağlayıcı değildir; çalışan yan
// panelden tek tıkla uygular ya da başka şube seçer.

import { getStockData } from "@/lib/stock-store";
import { stokEslesme, toBoy } from "@/lib/stock-search";
import { PERAKENDE_SABIT } from "@/lib/perakende-fiyat";
import type { IsKalem, Sube } from "./tur";

const PROFIL_VARSAYILAN_MM = 40; // dış ölçü bilinmiyorsa profil genişliği tahmini (kenar başına)

/** Kalem için gereken profil metresi: dış çevre (m) × adet + fire. */
export function gerekliMetre(k: IsKalem): number {
  const dis = k.disMm || (k.eserMm ? { w: k.eserMm.w + 2 * PROFIL_VARSAYILAN_MM, h: k.eserMm.h + 2 * PROFIL_VARSAYILAN_MM } : null);
  if (!dis) return 0;
  const cevre = (2 * (dis.w + dis.h)) / 1000;
  return Math.round((cevre + PERAKENDE_SABIT.FIRE_M) * Math.max(1, k.adet || 1) * 100) / 100;
}

export interface SubeOneri {
  sube: Sube;
  neden: string;
  detay: { sku: string; gerekliM: number; ankaraM: number; istanbulM: number; belirsiz: boolean; bulunamadi: boolean }[];
}

function sehirIpucu(sehir: string, adres: string): Sube {
  const t = `${sehir} ${adres}`.toLocaleLowerCase("tr-TR");
  if (/\bankara\b/.test(t)) return "ankara";
  if (/\bistanbul\b|\bİstanbul\b|\bistanbul\b/i.test(t) || t.includes("istanbul") || t.includes("i̇stanbul")) return "istanbul";
  return "";
}

/**
 * Öneri kuralı: müşteri Ankara/İstanbul'daysa o şube tercih edilir; tercih
 * edilen şubede her kalem için yeterli profil varsa o, yoksa diğer şubede
 * yeterliyse o; hiçbirinde yoksa tercih edilen (uyarıyla). Şehir yoksa stoğu
 * yeten şube, ikisi de yetiyorsa Ankara (merkez).
 */
export async function subeOner(kalemler: IsKalem[], sehir = "", adres = ""): Promise<SubeOneri> {
  const tercih = sehirIpucu(sehir, adres);
  let items: Awaited<ReturnType<typeof getStockData>>["items"] = [];
  try { items = (await getStockData()).items; } catch { items = []; }
  const detay: SubeOneri["detay"] = [];
  let ankaraYeter = true, istanbulYeter = true, bilinen = 0;
  for (const k of kalemler) {
    if (!k.sku) continue;
    const gerekli = gerekliMetre(k);
    const e = stokEslesme(items, k.sku);
    const belirsiz = !!e && !e.tam && e.aday > 1;
    const bulunamadi = !e;
    const ankaraM = e && !belirsiz ? e.item.ankaraMt : 0;
    const istanbulM = e && !belirsiz ? e.item.istanbulMt : 0;
    detay.push({ sku: k.sku, gerekliM: gerekli, ankaraM, istanbulM, belirsiz, bulunamadi });
    if (bulunamadi || belirsiz) continue;
    bilinen++;
    if (ankaraM < gerekli) ankaraYeter = false;
    if (istanbulM < gerekli) istanbulYeter = false;
  }
  const ozet = detay
    .map((d) => d.bulunamadi ? `${d.sku}: stokta bulunamadı` : d.belirsiz ? `${d.sku}: kod belirsiz` : `${d.sku}: Ankara ${toBoy(d.ankaraM)} boy · İstanbul ${toBoy(d.istanbulM)} boy (gerekli ~${d.gerekliM.toLocaleString("tr-TR")} m)`)
    .join("; ");
  const sehirTxt = tercih ? `müşteri ${tercih === "ankara" ? "Ankara" : "İstanbul"}` : "şehir bilinmiyor";
  if (!bilinen) {
    const sube: Sube = tercih || "ankara";
    return { sube, neden: `Stok eşleşmedi (${ozet || "kalem yok"}); ${sehirTxt} → ${sube === "ankara" ? "Ankara" : "İstanbul"} (varsayılan)`, detay };
  }
  let sube: Sube;
  let neden: string;
  if (tercih) {
    const tercihYeter = tercih === "ankara" ? ankaraYeter : istanbulYeter;
    const digerYeter = tercih === "ankara" ? istanbulYeter : ankaraYeter;
    const diger: Sube = tercih === "ankara" ? "istanbul" : "ankara";
    if (tercihYeter) { sube = tercih; neden = `${sehirTxt}, ${sube === "ankara" ? "Ankara" : "İstanbul"} stoğu yeterli`; }
    else if (digerYeter) { sube = diger; neden = `${sehirTxt} ama stok yalnızca ${diger === "ankara" ? "Ankara" : "İstanbul"}'da yeterli`; }
    else { sube = tercih; neden = `${sehirTxt}; iki depoda da stok yetersiz görünüyor, kontrol edin`; }
  } else if (ankaraYeter && istanbulYeter) { sube = "ankara"; neden = "İki depoda da stok var; şehir bilinmiyor → Ankara (merkez)"; }
  else if (ankaraYeter) { sube = "ankara"; neden = "Stok yalnızca Ankara'da yeterli"; }
  else if (istanbulYeter) { sube = "istanbul"; neden = "Stok yalnızca İstanbul'da yeterli"; }
  else { sube = "ankara"; neden = "İki depoda da stok yetersiz görünüyor, kontrol edin"; }
  return { sube, neden: `${neden}. ${ozet}`, detay };
}
