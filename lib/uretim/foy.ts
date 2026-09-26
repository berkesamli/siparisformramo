// Üretim föyü (PDF) — takvimdeki iş için üretim föyü üretir ve Blob'a yazar.
//
// • perakende (mağaza) işi: föy zaten /api/perakende/orders/pdf'te var; burada
//   yeniden üretilmez, aynı sipariş kaydından üretilip akıtılır (tek tasarım).
// • online (ikas) ve elle/teklif işleri: kalemlerdeki `retail` verisinden
//   SavedRetailOrder benzeri bir kayıt kurulur, generateRetailPdf ile şube
//   rozeti + içerik/yön satırıyla üretilir, uretim/foy/<isId>.pdf'e yazılır.
// • yüklenen föy (ek): kullanıcı Claude'da ürettiği föyü yüklediyse foyYol o eke
//   işaret eder; yeniden üretim onu ezmez (kullanıcı "yeniden üret" derse ezer).

import { generateRetailPdf, type FoyEkstra } from "@/lib/retail-pdf";
import type { RetailItem, SavedRetailOrder } from "@/lib/retail-orders";
import { getRetailOrder } from "@/lib/retail-orders";
import { foyKaydet } from "./ek";
import { isGuncelle } from "./db";
import type { UretimIs } from "./tur";

/** İşin kalemlerinden föy için sipariş kaydı kurar (retail alanı olmayan kalem atlanır). */
export function isSiparisKaydi(is: UretimIs): SavedRetailOrder | null {
  const items: RetailItem[] = [];
  for (const k of is.kalemler) {
    if (!k.retail) continue;
    // Adet: föy diyagramı kalem başına çizilir; adet > 1 ise aynı kalem tekrar eklenmez (adet göstergede)
    items.push(k.retail);
  }
  if (!items.length) return null;
  const adet = is.kalemler.reduce((t, k) => t + (Number(k.adet) || 1), 0);
  const notlar = [is.notlar, adet > items.length ? `TOPLAM ${adet} ADET (kalem başına adet: ${is.kalemler.map((k) => `${k.sku}×${k.adet || 1}`).join(", ")})` : ""]
    .filter(Boolean).join(" — ");
  return {
    orderId: is.kaynakRef ? (is.kaynak === "online" ? `Online #${is.kaynakRef}` : is.kaynakRef) : is.id.toUpperCase(),
    dateKey: is.createdAt.slice(0, 10),
    createdAt: is.createdAt,
    updatedAt: is.updatedAt,
    status: "Beklemede",
    employee: is.olusturan,
    customerName: is.musteriAd || is.baslik,
    customerPhone: is.musteriTel,
    customerEmail: "",
    customerAddress: is.musteriAdres,
    branch: is.sube || undefined,
    usdRate: 0,
    deliveryDate: is.teslimTarih || "",
    notes: notlar.slice(0, 200),
    items,
    gross: is.tutar,
    discount: 0,
    total: is.tutar,
  };
}

export function isFoyEkstra(is: UretimIs): FoyEkstra {
  const icerik = [...new Set(is.kalemler.map((k) => k.icerik || "").filter(Boolean))].join(" / ");
  const yon = [...new Set(is.kalemler.map((k) => k.yon || "").filter(Boolean))].join(" / ");
  return {
    sube: is.sube,
    icerik: icerik || undefined,
    yon: yon || undefined,
    kaynak: is.kaynak === "online" ? `Online sipariş #${is.kaynakRef}` : is.kaynak === "teklif" ? `Teklif ${is.kaynakRef}` : undefined,
  };
}

/** Föy PDF'ini üretir (Blob'a yazmaz). Perakende işi kaynak siparişten üretir. */
export async function foyUret(is: UretimIs): Promise<Buffer | null> {
  if (is.kaynak === "perakende" && is.kaynakKey && is.kaynakRef) {
    const o = await getRetailOrder(is.kaynakKey, is.kaynakRef);
    if (o) return generateRetailPdf(o, { sube: is.sube || o.branch || "" });
  }
  const o = isSiparisKaydi(is);
  if (!o) return null;
  return generateRetailPdf(o, isFoyEkstra(is));
}

/** Föyü üretip Blob'a yazar ve işe işler; blob yoksa null (PDF yine akıtılabilir). */
export async function foyUretVeKaydet(is: UretimIs, by: string): Promise<{ pdf: Buffer; yol: string | null } | null> {
  const pdf = await foyUret(is);
  if (!pdf) return null;
  const yol = await foyKaydet(is.id, pdf);
  if (yol && yol !== is.foyYol) await isGuncelle(is.id, { foyYol: yol }, by);
  return { pdf, yol };
}
