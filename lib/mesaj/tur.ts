// Mesajlar (gelen kutusu) — ortak tipler.
// Kanal: whatsapp (Cloud API numarası), instagram (olga.cerceve DM), email (Gmail hesapları).

export type Kanal = "whatsapp" | "instagram" | "email";
export type KonusmaDurum = "acik" | "yanitlandi" | "kapali";
export type Yon = "gelen" | "giden";

export interface Ek {
  tur: "image" | "document" | "audio" | "video" | "sticker" | "other";
  ad: string;      // dosya adı / açıklama
  url?: string;    // /api/mesaj/ek?... (özel blob) ya da dış bağlantı
  mime?: string;
  boyut?: number;
}

export interface Konusma {
  id: string;
  kanal: Kanal;
  hesap: string;         // bizim taraf: WA phone id, IG hesap id, gmail adresi
  disKimlik: string;     // karşı taraf: 90… telefon, IGSID, e-posta adresi
  ad: string;            // karşı tarafın görünen adı
  baslik: string;        // e-posta konusu / IG kullanıcı adı
  musteriId: string | null;  // toptan (C…) ya da perakende (P…) müşteri kartı
  musteriTur: "toptan" | "perakende" | null;
  durum: KonusmaDurum;
  atanan: string | null; // çalışan kullanıcı adı
  sonMesajAt: string;    // ISO
  sonMesajOzet: string;
  sonGelenAt: string | null; // 24 saat penceresi
  okunmamis: number;
  meta: Record<string, unknown>; // kanal özel (gmail threadId/messageId, IG kullanıcı adı…)
  createdAt: string;
}

export interface Mesaj {
  id: string;
  konusmaId: string;
  yon: Yon;
  govde: string;
  /** E-posta: temizlenmiş HTML gövde (varsa arayüz bunu kum havuzlu çerçevede gösterir; düz metin yedeği govde'de). */
  html?: string;
  ekler: Ek[];
  disId: string | null;  // wamid / IG mid / e-posta Message-ID (tekrar önleme)
  gonderen: string;      // giden: çalışan adı; gelen: karşı taraf adı
  taslakAi: boolean;     // yapay zekâ taslağından gönderildi
  durum: string;         // giden: gonderildi | iletildi | okundu | hata; gelen: ""
  hata: string | null;
  at: string;
}

export interface KonusmaFiltre {
  kanal?: Kanal | "";
  hesap?: string;        // bizim taraf (örn. e-posta adresi) — hesapları ayırmak için
  durum?: KonusmaDurum | "";
  atanan?: string | "";  // "ben" çağıranın kullanıcı adıyla değiştirilir
  q?: string;
  limit?: number;
}

export const KANAL_ADI: Record<Kanal, string> = { whatsapp: "WhatsApp", instagram: "Instagram", email: "E-posta" };

/** Hesap adının kısa hâli (etiket için): "olgacercevee@gmail.com" → "olgacercevee". */
export function hesapKisa(hesap: string, max = 16): string {
  const yerel = String(hesap || "").split("@")[0] || String(hesap || "");
  return yerel.length > max ? yerel.slice(0, max - 1) + "…" : yerel;
}

/**
 * Hesap için harf rozeti: "olgacercevee@gmail.com" → "O". Aynı harfle başlayan başka hesap varsa ilk iki harf ("OL").
 * Listede uzun adres yerine renkli harf gösterilir; tam adres ipucunda ve konuşma başlığında durur.
 */
export function hesapHarf(hesap: string, hepsi: string[] = []): string {
  const yerel = (h: string) => (String(h || "").split("@")[0] || "").toLocaleUpperCase("tr-TR");
  const bu = yerel(hesap);
  if (!bu) return "?";
  const cakisan = hepsi.some((h) => h !== hesap && yerel(h)[0] === bu[0]);
  return cakisan ? bu.slice(0, 2) : bu[0];
}

/**
 * Liste önizlemesi: "Konu:" satırı, Gmail'in düz metindeki "[image: …]" / "[https://…]" kalıntıları ve
 * bağlantılar atılır, boşluklar toplanır. Geriye metin kalmazsa içeriğin türü söylenir (Görsel / Bağlantı).
 */
export function ozetTemizle(metin: string, max = 140): string {
  const ham = String(metin || "").replace(/^Konu: .*\n+/, "");
  // Kayıtlı özet 140 karakterde kesilmiş olabilir: kapanmamış "[image: …" / "[https://…" de atılır
  const s = ham
    .replace(/\[(image|cid):[^\]]*(\]|$)/gi, " ")
    .replace(/\[?<?https?:\/\/[^\s\]>]*[\]>]?/gi, " ")
    .replace(/\s+/g, " ")
    .replace(/^[\s|•·>*_=\[\]-]+|[\s\[\]<>]+$/g, "")
    .trim();
  if (!s) {
    if (/\[image:|\.(jpe?g|png|gif|webp)\b/i.test(ham)) return "Görsel";
    if (/https?:\/\//i.test(ham)) return "Bağlantı";
    return "";
  }
  return s.length > max ? s.slice(0, max - 1) + "…" : s;
}

/** Hangi kanallar ayarlı; whatsappSablon: 24 saat dışı / ilk mesaj için onaylı şablon tanımlı mı. */
export type KanalDurumu = Record<Kanal, boolean> & { whatsappSablon: boolean };

/** Konuşma penceresi: WhatsApp/Instagram'da müşterinin son mesajından itibaren 24 saat serbest yanıt. */
export function pencereAcik(k: Pick<Konusma, "kanal" | "sonGelenAt">, simdi = Date.now()): boolean {
  if (k.kanal === "email") return true;
  if (!k.sonGelenAt) return false;
  return simdi - new Date(k.sonGelenAt).getTime() < 24 * 3600_000;
}
