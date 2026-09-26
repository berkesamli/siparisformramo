// ikas (online) siparişlerini üretim takvimine taşır — "online" kaynağı.
// Mağaza (perakende) siparişlerinden AYRI tutulur.
//
// Akış: webhook ya da cron/açılış senkronu → sipariş API'den okunur → not
// çözümlenir → iş açılır/tazelenir (kaynaktanYaz) → yeni işe föy PDF'i üretilir
// → stoktan şube önerisi yazılır. Online iş plan tarihi olmadan "Yeni" olarak
// planlanmamış kuyruğa düşer; çalışan takvime sürükler.

import { ayarOku, ayarYaz, isBul, isBulKaynak, isGuncelle, kaynaktanYaz, type KaynakKayit } from "@/lib/uretim/db";
import { foyUretVeKaydet } from "@/lib/uretim/foy";
import { subeOner } from "@/lib/uretim/sube-oneri";
import type { IsDurum, Sube } from "@/lib/uretim/tur";
import { ikasConfigured, ikasSiparis, ikasSiparisListesi, ikasZaman, type IkasOrder } from "./client";
import { ikasSiparisKalemleri } from "./not";

/** ikas durumları → üretim durumu. İlk kayıtta: yeni (planlanmamış). */
export function ikasDurum(o: IkasOrder): IsDurum {
  const s = String(o.status || "").toUpperCase();
  const p = String(o.orderPackageStatus || "").toUpperCase();
  if (["CANCELLED", "REFUNDED", "REFUND_REQUEST_ACCEPTED"].includes(s) || p === "CANCELLED" || p === "REFUNDED") return "iptal";
  if (p === "DELIVERED" || p === "FULFILLED" || p === "PARTIALLY_DELIVERED") return "teslim";
  if (p === "READY_FOR_SHIPMENT" || p === "READY_FOR_PICK_UP" || p === "PARTIALLY_READY_FOR_SHIPMENT" || p === "PARTIALLY_FULFILLED") return "hazir";
  return "yeni";
}

function adSoyad(o: IkasOrder): string {
  const a = o.shippingAddress;
  const s = [a?.firstName, a?.lastName].filter(Boolean).join(" ").trim();
  return s || o.customer?.fullName || [o.customer?.firstName, o.customer?.lastName].filter(Boolean).join(" ").trim() || `Sipariş #${o.orderNumber || ""}`;
}

function adres(o: IkasOrder): { adres: string; sehir: string } {
  const a = o.shippingAddress || o.billingAddress;
  if (!a) return { adres: "", sehir: "" };
  const sehir = a.city?.name || "";
  const parca = [a.addressLine1, a.addressLine2, a.district?.name, sehir].filter(Boolean).map((x) => String(x).trim());
  return { adres: parca.join(" ").replace(/\s+/g, " ").slice(0, 240), sehir };
}

/**
 * Zaman çizelgesi notları ikas API'sinden OKUNAMAZ (yalnızca yazılabilir). Hesaplayıcı
 * sunucusu aynı metni /api/ikas/not ucumuza da gönderirse burada (uretim_ayar
 * "ikasnot:<OZL>" / "ikasnot:no:<siparişNo>") bulunur ve kalem çözümüne katılır.
 */
export async function kayitliNotlar(o: IkasOrder): Promise<string> {
  const anahtarlar = new Set<string>();
  for (const s of o.orderLineItems || []) if (s.variant?.sku) anahtarlar.add(`ikasnot:${String(s.variant.sku).toUpperCase()}`);
  if (o.orderNumber) anahtarlar.add(`ikasnot:no:${String(o.orderNumber).replace(/^#/, "")}`);
  const parcalar: string[] = [];
  for (const a of anahtarlar) {
    try { const r = await ayarOku(a); if (r?.deger) parcalar.push(r.deger); } catch { /* not yoksa geç */ }
  }
  return parcalar.join("\n");
}

/** Siparişi kayda çevirir (kalem çözümü + şube önerisi). */
export async function ikasKaydi(o: IkasOrder, ekMetin = ""): Promise<KaynakKayit> {
  const kayitli = await kayitliNotlar(o);
  const { kalemler, notMetni, eksik } = ikasSiparisKalemleri(o, [ekMetin, kayitli].filter(Boolean).join("\n"));
  const { adres: adr, sehir } = adres(o);
  let subeOneri: Sube = "", subeOneriNeden = "";
  try {
    const on = await subeOner(kalemler, sehir, adr);
    subeOneri = on.sube; subeOneriNeden = on.neden;
  } catch (e) {
    console.warn("Şube önerisi hesaplanamadı:", (e as Error)?.message);
  }
  const notSerbest = (o.note || "").trim();
  // Hesaplayıcı bloğu notta duruyorsa "not" olarak tekrar gösterme
  const notlar = /cerceve siparis detayi/i.test(notSerbest) ? "" : notSerbest.slice(0, 2000);
  const kargo = (o.orderPackages || []).map((p) => p.trackingInfo).filter((t) => t && (t.trackingNumber || t.cargoCompany))
    .map((t) => `${t!.cargoCompany || ""} ${t!.trackingNumber || ""}`.trim());
  return {
    kaynak: "online",
    kaynakRef: String(o.orderNumber || o.id),
    kaynakKey: o.id,
    kaynakAt: ikasZaman(o.updatedAt) || ikasZaman(o.orderedAt) || new Date().toISOString(),
    kaynakDurum: `${o.status || ""}/${o.orderPackageStatus || ""}`,
    durum: ikasDurum(o),
    musteriAd: adSoyad(o),
    musteriTel: String(o.shippingAddress?.phone || o.customer?.phone || "").trim(),
    musteriAdres: adr,
    musteriSehir: sehir,
    subeOneri, subeOneriNeden,
    teslimTarih: ikasZaman(o.dueDate)?.slice(0, 10) || null,
    kalemler,
    tutar: Number(o.totalFinalPrice ?? o.totalPrice) || 0,
    notlar,
    ham: {
      ikasId: o.id, orderNumber: o.orderNumber, status: o.status, orderPackageStatus: o.orderPackageStatus,
      orderPaymentStatus: o.orderPaymentStatus, orderedAt: ikasZaman(o.orderedAt), salesChannel: o.salesChannel?.name || o.salesChannelId || "",
      eposta: o.customer?.email || "", notMetni: notMetni.slice(0, 8000), eksikKalem: eksik, kargo, kaynakNot: notlar,
    },
  };
}

export interface IsleSonuc { isId: string; yeni: boolean; guncellendi: boolean; foy: boolean }

/** Tek siparişi işler (webhook ve senkron ortak yolu). Yeni işe föy üretilir. */
export async function ikasSiparisIsle(o: IkasOrder, by: string, ekMetin = ""): Promise<IsleSonuc> {
  if (String(o.status || "").toUpperCase() === "DRAFT") throw new Error("Taslak sipariş atlandı.");
  const kayit = await ikasKaydi(o, ekMetin);
  const r = await kaynaktanYaz(kayit, by);
  let foy = false;
  const kalemliMi = kayit.kalemler.some((k) => k.retail);
  if (kalemliMi && (r.yeni || (r.guncellendi && !r.is.foyYol))) {
    try {
      const f = await foyUretVeKaydet(r.is, by);
      foy = Boolean(f?.yol);
    } catch (e) {
      console.warn("Online föy üretilemedi:", (e as Error)?.message);
    }
  }
  // Şube önerisi zamanla değişebilir (stok tazelenir): var olan işte öneri güncellenir, seçili şube ellenmez
  if (!r.yeni && r.guncellendi && kayit.subeOneri && kayit.subeOneri !== r.is.subeOneri) {
    await isGuncelle(r.is.id, { subeOneri: kayit.subeOneri, subeOneriNeden: kayit.subeOneriNeden || "" }, by, true);
  }
  return { isId: r.is.id, yeni: r.yeni, guncellendi: r.guncellendi, foy };
}

export interface IkasSenkSonuc { eklenen: number; guncellenen: number; okunan: number; hata?: string; atlandi?: string }

const AYAR_SON = "ikas_senk_at";

/**
 * Cron / açılış senkronu: son senkrondan (15 dk payla) beri güncellenen siparişler;
 * ilk çalışmada son 30 gün (tam=true: 90 gün). Sayfa sayfa okur, DRAFT'ı atlar.
 */
export async function ikasSenk(by = "ikas", secenek: { tam?: boolean } = {}): Promise<IkasSenkSonuc> {
  if (!ikasConfigured()) return { eklenen: 0, guncellenen: 0, okunan: 0, atlandi: "IKAS_CLIENT_ID / IKAS_CLIENT_SECRET tanımlı değil." };
  const baslangic = Date.now();
  const son = await ayarOku(AYAR_SON);
  const gun = 86_400_000;
  const since = secenek.tam || !son ? baslangic - (secenek.tam ? 90 : 30) * gun : Date.parse(son.deger) - 15 * 60_000;
  let eklenen = 0, guncellenen = 0, okunan = 0;
  let hata: string | undefined;
  try {
    // updatedAt süzgeci şemada yoksa orderedAt ile dener
    let filtre: "updatedAt" | "orderedAt" = "updatedAt";
    for (let page = 1; page <= 40; page++) {
      let r;
      try {
        r = await ikasSiparisListesi(filtre === "updatedAt" ? { updatedAtGte: since, page, limit: 100 } : { orderedAtGte: since, page, limit: 100 });
      } catch (e) {
        if (filtre === "updatedAt" && /updatedAt/i.test(String((e as Error)?.message))) { filtre = "orderedAt"; page--; continue; }
        throw e;
      }
      for (const o of r.data) {
        if (String(o.status || "").toUpperCase() === "DRAFT") continue;
        okunan++;
        try {
          const s = await ikasSiparisIsle(o, by);
          if (s.yeni) eklenen++; else if (s.guncellendi) guncellenen++;
        } catch (e) {
          console.error(`ikas siparişi işlenemedi (#${o.orderNumber}):`, (e as Error)?.message);
        }
      }
      if (!r.hasNext) break;
    }
    await ayarYaz(AYAR_SON, new Date(baslangic).toISOString());
  } catch (e) {
    hata = (e as Error)?.message || String(e);
    console.error("ikas senkronu:", hata);
  }
  return { eklenen, guncellenen, okunan, hata };
}

/** Son senkron `dk` dakikadan eskiyse çalıştırır (takvim açılışı). */
export async function ikasSenkGerekirse(by: string, dk = 10): Promise<IkasSenkSonuc | null> {
  if (!ikasConfigured()) return null;
  const son = await ayarOku(AYAR_SON);
  if (son && Date.now() - Date.parse(son.deger) < dk * 60_000) return null;
  return ikasSenk(by);
}

/** Webhook: kimliği verilen siparişi API'den okuyup işler; API yoksa yükteki veriyle çalışır. */
export async function ikasWebhookIsle(yuk: Partial<IkasOrder> & { id?: string }, by = "ikas-webhook"): Promise<IsleSonuc | null> {
  let o: IkasOrder | null = null;
  if (ikasConfigured() && yuk.id) {
    try { o = await ikasSiparis(yuk.id); } catch (e) { console.warn("ikas siparişi okunamadı, webhook verisi kullanılacak:", (e as Error)?.message); }
  }
  if (!o) {
    if (!yuk.id || !Array.isArray(yuk.orderLineItems)) return null;
    o = yuk as IkasOrder;
  }
  return ikasSiparisIsle(o, by);
}

/** Sipariş numarasıyla tek seferlik içe aktarma (test / elle / not geldiğinde). Var olan iş zorla tazelenir. */
export async function ikasSiparisNoIsle(orderNumber: string, by: string): Promise<IsleSonuc | null> {
  const { ikasSiparisNo } = await import("./client");
  const o = await ikasSiparisNo(orderNumber);
  if (!o) return null;
  const mevcut = await isBulKaynak("online", String(o.orderNumber || o.id));
  if (mevcut) o.updatedAt = new Date().toISOString(); // kaynaktanYaz "daha yeni" görsün, kalemler yeniden çözülsün
  const r = await ikasSiparisIsle(o, by);
  if (mevcut && r.guncellendi && !r.foy) {
    try { const is = await isBul(r.isId); if (is && is.kalemler.some((k) => k.retail)) await foyUretVeKaydet(is, by); } catch { /* föy sonra da üretilebilir */ }
  }
  return r;
}
