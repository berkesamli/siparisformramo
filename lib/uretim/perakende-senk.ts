// Mağaza (perakende) siparişlerini üretim takvimine taşır.
//
// Perakende siparişleri Vercel Blob'da durur (lib/retail-orders). Bu modül
// son aylardaki siparişleri okur ve her birini "perakende" kaynaklı bir işe
// çevirir: teslim tarihi varsa iş o güne planlanır, yoksa planlanmamış
// kuyruğa düşer. Durum eşlemesi iki yönlüdür:
//   perakende → iş : Beklemede→planlandı, Hazırlanıyor→üretimde, Hazır→hazır,
//                    Teslim Edildi→teslim, İptal→iptal (kaynaktanYaz kuralları)
//   iş → perakende : takvimden durum değişince sipariş de güncellenir
//                    (perakendeDurumYaz).
//
// Online (ikas) siparişleri AYRI kaynaktır (lib/ikas); burada karışmaz.

import {
  getRetailOrder, listAllRetailOrders, readRetailIndexOrRebuild, saveRetailOrder, type RetailItem, type SavedRetailOrder,
} from "@/lib/retail-orders";
import { blobConfigured } from "@/lib/orders";
import { RETAIL_STATUSES, type RetailStatus } from "@/data/perakende";
import { kaynaktanYaz, ayarOku, ayarYaz, isBul, type KaynakKayit } from "./db";
import { IS_DURUM_PERAKENDE, PERAKENDE_DURUM_ESLE, TARIH_RE, olcuCm, type IsDurum, type IsKalem, type Sube } from "./tur";

const mmOlcu = (v: number, u: "cm" | "mm") => (u === "cm" ? v * 10 : v);

/** Perakende kalemini takvim kalemine çevirir (retail alanı föy için korunur). */
export function perakendeKalem(it: RetailItem): IsKalem {
  const w = mmOlcu(Number(it.artWidth) || 0, it.artWidthUnit);
  const h = mmOlcu(Number(it.artHeight) || 0, it.artHeightUnit);
  const paspartu = it.matType && it.matType !== "Paspartu Yok"
    ? `${it.matType}${it.doubleMat ? " (çift)" : ""}${it.matTop ? ` ${it.matTop}/${it.matRight}/${it.matBottom}/${it.matLeft} mm` : ""}`
    : "Yok";
  const parcalar = [`Eser ${olcuCm({ w, h })}`];
  if (it.kasa) parcalar.push("Kasa (kanvas)");
  parcalar.push(paspartu === "Yok" ? "Paspartu yok" : `Paspartu: ${paspartu}`);
  if (it.glassType && it.glassType !== "Cam Yok") parcalar.push(it.glassType);
  if (it.printType && it.printType !== "Baskı Yok") parcalar.push(`Baskı: ${it.printType}`);
  return {
    sku: it.frameCode || "OZEL",
    adet: 1,
    ozet: parcalar.join(" · "),
    eserMm: { w, h },
    paspartu,
    cam: it.glassType || "",
    fiyat: Number(it.itemTotal) || 0,
    retail: it,
  };
}

function siparisKaydi(o: SavedRetailOrder): KaynakKayit {
  const durum: IsDurum = PERAKENDE_DURUM_ESLE[o.status] || "planlandi";
  const teslim = TARIH_RE.test(o.deliveryDate || "") ? o.deliveryDate : null;
  return {
    kaynak: "perakende",
    kaynakRef: o.orderId,
    kaynakKey: o.dateKey,
    kaynakAt: o.updatedAt || o.createdAt,
    kaynakDurum: o.status,
    durum,
    musteriAd: o.customerName,
    musteriTel: o.customerPhone,
    musteriAdres: o.customerAddress || "",
    sube: (o.branch || "") as Sube,
    teslimTarih: teslim,
    kalemler: (o.items || []).map(perakendeKalem),
    tutar: Number(o.total) || 0,
    notlar: o.notes || "",
    ham: { kaynakNot: o.notes || "", employee: o.employee, payment: o.payment || "", paidAmount: o.paidAmount || 0, dateKey: o.dateKey },
  };
}

export interface SenkSonuc { eklenen: number; guncellenen: number; okunan: number; atlandi?: string }

/**
 * Son `sonAy` ayın perakende siparişlerini işlere yazar. İndeksten hangi
 * siparişlerin yeni/değişmiş olduğu anlaşılır; yalnızca onların tam kaydı
 * okunur (her siparişi tek tek okumak N+1 olurdu).
 */
export async function perakendeSenk(by = "senk", sonAy = 3, zorla = false): Promise<SenkSonuc> {
  if (!blobConfigured()) return { eklenen: 0, guncellenen: 0, okunan: 0, atlandi: "Blob depo bağlı değil." };
  const idx = await readRetailIndexOrRebuild(sonAy);
  let eklenen = 0, guncellenen = 0, okunan = 0;
  const { isBulKaynak } = await import("./db");
  for (const e of idx) {
    // Hızlı eleme: iş var ve kaynak güncellenmemişse dokunma (indeks updatedAt taşımaz; createdAt + durum kıyası)
    const mevcut = zorla ? null : await isBulKaynak("perakende", e.orderId);
    if (mevcut && String(mevcut.ham.kaynakDurum || "") === e.status && (mevcut.teslimTarih || "") === (TARIH_RE.test(e.deliveryDate || "") ? e.deliveryDate : "")) continue;
    const o = await getRetailOrder(e.dateKey, e.orderId);
    if (!o) continue;
    okunan++;
    const r = await kaynaktanYaz(siparisKaydi(o), by);
    if (r.yeni) eklenen++; else if (r.guncellendi) guncellenen++;
  }
  await ayarYaz("perakende_senk_at", new Date().toISOString());
  return { eklenen, guncellenen, okunan };
}

/** Son senkron üzerinden `dk` dakika geçtiyse senkronu çalıştırır (takvim açılışı). */
export async function perakendeSenkGerekirse(by: string, dk = 3): Promise<SenkSonuc | null> {
  const son = await ayarOku("perakende_senk_at");
  if (son && Date.now() - Date.parse(son.deger) < dk * 60_000) return null;
  return perakendeSenk(by);
}

/** Tam tarama (indeks bozuksa): bütün perakende siparişleri. */
export async function perakendeTamSenk(by = "senk"): Promise<SenkSonuc> {
  if (!blobConfigured()) return { eklenen: 0, guncellenen: 0, okunan: 0, atlandi: "Blob depo bağlı değil." };
  const hepsi = await listAllRetailOrders();
  let eklenen = 0, guncellenen = 0;
  for (const o of hepsi) {
    const r = await kaynaktanYaz(siparisKaydi(o), by);
    if (r.yeni) eklenen++; else if (r.guncellendi) guncellenen++;
  }
  await ayarYaz("perakende_senk_at", new Date().toISOString());
  return { eklenen, guncellenen, okunan: hepsi.length };
}

/**
 * Takvimde durum değişince perakende siparişini de günceller (iş → sipariş).
 * İşin ham.kaynakDurum'u da yeni değerle yazılır ki bir sonraki senkron bunu
 * "kaynak değişti" sanıp durumu geri almasın. Sessizce başarısız olabilir.
 */
export async function perakendeDurumYaz(isId: string, durum: IsDurum): Promise<boolean> {
  const is = await isBul(isId);
  if (!is || is.kaynak !== "perakende" || !is.kaynakRef || !is.kaynakKey) return false;
  const hedef = IS_DURUM_PERAKENDE[durum];
  if (!hedef || !RETAIL_STATUSES.includes(hedef as RetailStatus)) return false;
  const o = await getRetailOrder(is.kaynakKey, is.kaynakRef);
  if (!o) return false;
  if (o.status === hedef) return true;
  o.status = hedef as RetailStatus;
  o.updatedAt = new Date().toISOString();
  const ok = await saveRetailOrder(o);
  if (ok) {
    const { isGuncelle } = await import("./db");
    await isGuncelle(isId, { ham: { ...is.ham, kaynakDurum: hedef }, kaynakAt: o.updatedAt }, "senk", true);
  }
  return ok;
}
