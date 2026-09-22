// Tek kapı: konuşmanın kanalına göre yanıtı gönderir ve mesaj olarak kaydeder;
// WhatsApp'ta bizim başlattığımız yeni konuşma da buradan açılır.

import { normalizeWaNumber } from "@/lib/whatsapp-pdf";
import { konusma, konusmaBulVeyaOlustur, konusmaGuncelle, mesajEkle } from "./db";
import { gmailConfigured, gmailGonder } from "./gmail";
import { instagramConfigured, instagramGonder } from "./instagram";
import { musteriEsle } from "./musteri-esle";
import { sablonParam, whatsappConfigured, whatsappGonder, whatsappSablonHazir } from "./whatsapp";
import type { KanalDurumu, Konusma, Mesaj } from "./tur";

export function kanalDurumu(): KanalDurumu {
  return { whatsapp: whatsappConfigured(), instagram: instagramConfigured(), email: gmailConfigured(), whatsappSablon: whatsappSablonHazir() };
}

export interface GonderimSonucu { mesaj: Mesaj; yontem: "serbest" | "sablon" | "eposta" | "instagram"; sablon?: string }

interface Kullanici { username: string; name: string }

async function gonderVeKaydet(k: Konusma, govde: string, kullanici: Kullanici, taslakAi: boolean): Promise<GonderimSonucu> {
  let disId: string | null = null;
  let yontem: GonderimSonucu["yontem"];
  let sablon: string | undefined;
  if (k.kanal === "whatsapp") {
    const g = await whatsappGonder(k, govde);
    disId = g.disId; yontem = g.yontem; sablon = g.sablon;
  } else if (k.kanal === "instagram") {
    disId = (await instagramGonder(k, govde)).disId; yontem = "instagram";
  } else {
    const r = await gmailGonder(k, govde);
    disId = r.disId; yontem = "eposta";
    await konusmaGuncelle(k.id, { meta: { ...k.meta, references: r.references } });
  }
  // Şablonla gidince satır sonları boşluğa döner; kayıt, müşterinin gerçekten gördüğü metin olsun.
  const kayit = yontem === "sablon" ? sablonParam(govde) : govde;
  const { mesaj } = await mesajEkle({ konusmaId: k.id, yon: "giden", govde: kayit, disId, gonderen: kullanici.name, taslakAi, durum: "gonderildi" });
  return { mesaj, yontem, sablon };
}

export async function yanitGonder(konusmaId: string, metin: string, kullanici: Kullanici, taslakAi = false): Promise<GonderimSonucu> {
  const k = await konusma(konusmaId);
  if (!k) throw new Error("Konuşma bulunamadı.");
  const govde = String(metin || "").trim();
  if (!govde) throw new Error("Boş mesaj gönderilemez.");
  if (govde.length > 4000 && k.kanal !== "email") throw new Error("Mesaj çok uzun (en fazla 4000 karakter).");
  return gonderVeKaydet(k, govde, kullanici, taslakAi);
}

/**
 * Bizim başlattığımız WhatsApp mesajı: numaraya konuşma açılır (varsa bulunur), müşteri kartı
 * bağlanır ve mesaj gönderilir — müşteri son 24 saatte yazmadıysa onaylı şablonla.
 */
export async function yeniWhatsappMesaji(g: {
  telefon: string; ad?: string; metin: string; kullanici: Kullanici;
  musteriId?: string | null; musteriTur?: "toptan" | "perakende" | null; taslakAi?: boolean;
}): Promise<GonderimSonucu & { konusma: Konusma }> {
  if (!whatsappConfigured()) throw new Error("WhatsApp Cloud API ayarlı değil.");
  const to = normalizeWaNumber(g.telefon);
  if (!to) throw new Error("Telefon numarası geçersiz (örn. 0532 111 22 33).");
  const govde = String(g.metin || "").trim();
  if (!govde) throw new Error("Boş mesaj gönderilemez.");
  if (govde.length > 1000) throw new Error("İlk mesaj en fazla 1000 karakter olabilir (şablon sınırı).");
  const hesap = (process.env.WHATSAPP_PHONE_ID || "").trim();
  let k = await konusmaBulVeyaOlustur({ kanal: "whatsapp", hesap, disKimlik: to, ad: g.ad?.trim() || undefined, meta: { telefon: to } });
  if (!k.ad) k = (await konusmaGuncelle(k.id, { ad: `+${to}` })) || k;
  if (g.musteriId) {
    k = (await konusmaGuncelle(k.id, { musteriId: g.musteriId, musteriTur: g.musteriTur || (g.musteriId.startsWith("P") ? "perakende" : "toptan") })) || k;
  } else if (!k.musteriId) {
    const e = await musteriEsle({ telefon: to }).catch(() => null);
    if (e) k = (await konusmaGuncelle(k.id, { musteriId: e.musteriId, musteriTur: e.musteriTur, ad: k.ad.startsWith("+") ? e.ad : k.ad })) || k;
  }
  const r = await gonderVeKaydet(k, govde, g.kullanici, Boolean(g.taslakAi));
  return { ...r, konusma: (await konusma(k.id)) || k };
}
