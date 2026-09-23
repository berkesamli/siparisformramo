// WhatsApp Cloud API bağlayıcısı (gelen kutusu).
//   gelenWhatsapp(value)     — webhook "messages" alanını konuşma + mesaja çevirir
//   whatsappDurumlar(statuses) — giden mesajların teslim/okundu durumunu işler
//   whatsappGonder(konusma, metin) — serbest metin yanıtı (24 saat penceresi içinde)
//
// Aynı Meta uygulaması ve WHATSAPP_TOKEN / WHATSAPP_PHONE_ID kullanılır.
// Medya: Graph'tan indirilip özel Blob'a yazılır (bkz. ek.ts).

import { normalizeWaNumber, whatsappConfigured } from "@/lib/whatsapp-pdf";
import { dbConfigured, konusmaBulVeyaOlustur, konusmaGuncelle, mesajDurumDisId, mesajEkle, senkOku, senkYaz } from "./db";
import { ekIndir, ekKaydet, ekTuru, ekUrl, mimeUzanti } from "./ek";
import { musteriEsle } from "./musteri-esle";
import { pencereAcik, type Ek, type Konusma } from "./tur";

const GRAPH = "https://graph.facebook.com/v20.0";

export { whatsappConfigured };

interface WaMedya { id?: string; mime_type?: string; sha256?: string; caption?: string; filename?: string; animated?: boolean }
export interface WaMesaj {
  from: string;
  id: string;
  timestamp?: string;
  type?: string;
  text?: { body?: string };
  image?: WaMedya; video?: WaMedya; audio?: WaMedya; document?: WaMedya; sticker?: WaMedya;
  location?: { latitude?: number; longitude?: number; name?: string; address?: string };
  contacts?: { name?: { formatted_name?: string }; phones?: { phone?: string; wa_id?: string }[] }[];
  reaction?: { emoji?: string; message_id?: string };
  button?: { text?: string; payload?: string };
  interactive?: { type?: string; button_reply?: { title?: string }; list_reply?: { title?: string; description?: string } };
  context?: { from?: string; id?: string };
  errors?: { code?: number; title?: string; message?: string }[];
}
export interface WaValue {
  messaging_product?: string;
  metadata?: { display_phone_number?: string; phone_number_id?: string };
  contacts?: { profile?: { name?: string }; wa_id?: string }[];
  messages?: WaMesaj[];
}
export interface WaDurum {
  id?: string;
  status?: string;
  recipient_id?: string;
  errors?: { code?: number; title?: string; message?: string; error_data?: { details?: string } }[];
}

const token = () => (process.env.WHATSAPP_TOKEN || "").trim();
const phoneId = () => (process.env.WHATSAPP_PHONE_ID || "").trim();
const wabaId = () => (process.env.WHATSAPP_WABA_ID || "").trim();
const sablonDili = () => (process.env.WHATSAPP_TEMPLATE_DIL || "tr").trim();

/**
 * Bizim başlattığımız / 24 saat penceresi dışındaki mesajlar için Meta onaylı şablon(lar):
 * WHATSAPP_TEMPLATE_SERBEST="genel_mesaj,genel_mesaj_v2" (sırayla denenir). Gövdesinde
 * {{1}} = mesaj metni; iki değişkenliyse {{1}} = müşteri adı, {{2}} = metin.
 */
export function serbestSablonAdlari(): string[] {
  const out: string[] = [];
  for (const p of (process.env.WHATSAPP_TEMPLATE_SERBEST || "").split(/[,;]+/)) {
    const ad = p.trim();
    if (ad && !out.includes(ad)) out.push(ad);
  }
  return out;
}
/**
 * Varsayılan şablon: Ayarlar kartındaki "Şablonu oluştur" bununla Meta'ya başvurur; Meta onaylayınca
 * (durum APPROVED) env'de ad yazmaya gerek kalmadan kendiliğinden kullanılır.
 * {{1}} = müşteri adı, {{2}} = mesaj metni.
 */
export const SABLON_VARSAYILAN = "genel_mesaj";
export const SABLON_GOVDE = "Merhaba {{1}}, Olga Çerçeve'den yazıyoruz:\n\n{{2}}\n\nSorularınız için bu numaradan bize yazabilirsiniz.";
const SABLON_ORNEK = ["Ayşe Hanım", "Siparişiniz hazırlandı; teslimat için uygun olduğunuz günü yazabilir misiniz?"];

export interface SablonBilgi { ad: string; durum: string; kategori: string; dil: string; degisken: number; red?: string; id?: string }

let sablonOnbellek: { at: number; liste: SablonBilgi[] } | null = null;
const SABLON_ONBELLEK_MS = 10 * 60_000;
/** Testler için: şablon önbelleğini sıfırla. */
export function sablonOnbellekSifirla() { sablonOnbellek = null; }

function degiskenSayisi(components: { type?: string; text?: string }[] | undefined): number {
  const govde = (components || []).find((c) => (c.type || "").toUpperCase() === "BODY")?.text || "";
  const n = new Set((govde.match(/\{\{\s*(\d+)\s*\}\}/g) || []).map((x) => x.replace(/\D/g, "")));
  return n.size;
}

/**
 * WhatsApp hesabındaki (WABA) mesaj şablonları — ad, onay durumu, kategori, dil, değişken sayısı.
 * 10 dk bellekte ve veri tabanında (mesaj_senk "wa:sablon") önbelleklenir; zorla=true Meta'dan tazeler.
 */
export async function sablonListesi(zorla = false): Promise<{ ok: boolean; liste: SablonBilgi[]; hata?: string }> {
  if (!wabaId() || !token()) return { ok: false, liste: [], hata: "Şablonlar için WHATSAPP_WABA_ID ve WHATSAPP_TOKEN gerekli." };
  if (!zorla && sablonOnbellek && Date.now() - sablonOnbellek.at < SABLON_ONBELLEK_MS) return { ok: true, liste: sablonOnbellek.liste };
  if (!zorla && dbConfigured()) {
    try {
      const s = await senkOku("wa:sablon");
      if (s && Date.now() - new Date(s.at).getTime() < SABLON_ONBELLEK_MS) {
        const liste = JSON.parse(s.deger) as SablonBilgi[];
        sablonOnbellek = { at: new Date(s.at).getTime(), liste };
        return { ok: true, liste };
      }
    } catch { /* Meta'dan okunur */ }
  }
  try {
    const r = await fetch(`${GRAPH}/${wabaId()}/message_templates?fields=id,name,status,category,language,components,rejected_reason&limit=200`, { headers: { Authorization: `Bearer ${token()}` }, signal: AbortSignal.timeout(10_000) });
    const j = (await r.json().catch(() => ({}))) as { data?: { id?: string; name?: string; status?: string; category?: string; language?: string; components?: { type?: string; text?: string }[]; rejected_reason?: string }[]; error?: { message?: string; code?: number } };
    if (!r.ok) return { ok: false, liste: [], hata: `${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    const liste: SablonBilgi[] = (j.data || []).map((d) => ({
      ad: d.name || "?", durum: d.status || "?", kategori: d.category || "?", dil: d.language || "?", degisken: degiskenSayisi(d.components), id: d.id,
      red: d.rejected_reason && d.rejected_reason !== "NONE" ? d.rejected_reason : undefined,
    }));
    sablonOnbellek = { at: Date.now(), liste };
    if (dbConfigured()) await senkYaz("wa:sablon", JSON.stringify(liste)).catch(() => undefined);
    return { ok: true, liste };
  } catch (e) {
    return { ok: false, liste: [], hata: (e as Error)?.message || String(e) };
  }
}

/** Şablon adı → bilgi (onaylı listeden; bilinmiyorsa null). */
export async function sablonBilgisi(ad: string): Promise<SablonBilgi | null> {
  const { liste } = await sablonListesi();
  return liste.find((s) => s.ad === ad && s.dil === sablonDili()) || liste.find((s) => s.ad === ad) || null;
}

/**
 * Bizim başlattığımız mesajlarda denenecek şablonlar, sırayla: env'dekiler, ardından Meta'da
 * ONAYLI görünen varsayılan şablon (genel_mesaj). Hiçbiri yoksa boş liste.
 */
export async function etkinSablonlar(): Promise<string[]> {
  const out = serbestSablonAdlari();
  if (!out.includes(SABLON_VARSAYILAN)) {
    const b = await sablonBilgisi(SABLON_VARSAYILAN).catch(() => null);
    if (b && b.durum === "APPROVED" && b.degisken >= 1 && b.degisken <= 2) out.push(SABLON_VARSAYILAN);
  }
  return out;
}

export async function whatsappSablonHazir(): Promise<boolean> {
  return whatsappConfigured() && (await etkinSablonlar()).length > 0;
}

/**
 * Varsayılan şablonu Meta'ya oluşturur (kategori UTILITY; Meta gerekli görürse kategoriyi kendisi
 * değiştirir). Onay çoğu zaman dakikalar içinde, bazen 24 saate kadar sürer; durum kartta görünür.
 */
export async function sablonOlustur(): Promise<{ ok: boolean; id?: string; durum?: string; kategori?: string; hata?: string }> {
  if (!wabaId() || !token()) return { ok: false, hata: "Şablon için WHATSAPP_WABA_ID ve WHATSAPP_TOKEN gerekli." };
  try {
    const r = await fetch(`${GRAPH}/${wabaId()}/message_templates`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token()}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        name: SABLON_VARSAYILAN, language: sablonDili(), category: "UTILITY", allow_category_change: true,
        components: [{ type: "BODY", text: SABLON_GOVDE, example: { body_text: [SABLON_ORNEK] } }],
      }),
      signal: AbortSignal.timeout(15_000),
    });
    const j = (await r.json().catch(() => ({}))) as { id?: string; status?: string; category?: string; error?: { message?: string; code?: number; error_user_msg?: string; error_user_title?: string } };
    if (!r.ok || !j.id) {
      const e = j.error;
      return { ok: false, hata: `${e?.error_user_title ? e.error_user_title + ": " : ""}${e?.error_user_msg || e?.message || `HTTP ${r.status}`}${e?.code ? ` (kod ${e.code})` : ""}` };
    }
    sablonOnbellek = null;
    if (dbConfigured()) await senkYaz("wa:sablon", "[]").catch(() => undefined);
    return { ok: true, id: j.id, durum: j.status, kategori: j.category };
  } catch (e) {
    return { ok: false, hata: (e as Error)?.message || String(e) };
  }
}

export const SABLON_METIN_AZAMI = 1000;

/** Şablon değişkeni: satır sonu/sekme yasak, 4+ ardışık boşluk yasak, en çok ~1000 karakter. */
export function sablonParam(s: string, max = SABLON_METIN_AZAMI): string {
  const t = String(s ?? "").replace(/[\r\n\t]+/g, " ").replace(/\s{2,}/g, " ").trim();
  return (t || "—").slice(0, max);
}

const HATA_IPUCU: Record<number, string> = {
  190: "token geçersiz ya da süresi dolmuş; WHATSAPP_TOKEN'ı yenileyin",
  131047: "24 saat penceresi kapalı; müşteri önce yazmalı ya da onaylı şablon (WHATSAPP_TEMPLATE_SERBEST) kullanılmalı",
  131026: "numara WhatsApp'ta kayıtlı olmayabilir",
  131030: "alıcı Meta'daki izin listesinde değil (uygulama geliştirme modunda)",
  132000: "şablondaki değişken sayısı uyuşmuyor",
  132001: "şablon Meta'da yok ya da henüz onaylanmadı",
  132015: "şablon Meta tarafından duraklatıldı",
  132016: "şablon devre dışı",
};

type GrafSonuc = { ok: true; disId: string } | { ok: false; hata: string; kod?: number };

async function grafGonder(hesap: string, body: Record<string, unknown>): Promise<GrafSonuc> {
  const r = await fetch(`${GRAPH}/${hesap || phoneId()}/messages`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token()}`, "Content-Type": "application/json" },
    body: JSON.stringify({ messaging_product: "whatsapp", recipient_type: "individual", ...body }),
    signal: AbortSignal.timeout(15_000),
  });
  const j = (await r.json().catch(() => ({}))) as { messages?: { id: string }[]; error?: { message?: string; code?: number; error_data?: { details?: string } } };
  if (r.ok && j.messages?.[0]?.id) return { ok: true, disId: j.messages[0].id };
  const e = j.error;
  const kod = typeof e?.code === "number" ? e.code : undefined;
  const ipucu = kod && HATA_IPUCU[kod] ? ` → ${HATA_IPUCU[kod]}` : "";
  return { ok: false, kod, hata: `${e?.message || `HTTP ${r.status}`}${kod ? ` (kod ${kod})` : ""}${e?.error_data?.details ? " — " + e.error_data.details : ""}${ipucu}` };
}

/** Jeton + numara kontrolü: Graph'tan görünen numara ve doğrulanmış ad (Ayarlar test kartı). */
export async function whatsappDurum(): Promise<{ ok: boolean; numara?: string; ad?: string; kalite?: string; uygulama?: { id: string; ad: string }; hata?: string }> {
  if (!whatsappConfigured()) return { ok: false, hata: "WHATSAPP_TOKEN / WHATSAPP_PHONE_ID tanımlı değil." };
  try {
    const basliklar = { Authorization: `Bearer ${token()}` };
    const [r, ra] = await Promise.all([
      fetch(`${GRAPH}/${phoneId()}?fields=display_phone_number,verified_name,quality_rating`, { headers: basliklar, signal: AbortSignal.timeout(10_000) }),
      // Jetonun ait olduğu Meta uygulaması — webhook o uygulamada ayarlanır.
      fetch(`${GRAPH}/app?fields=id,name`, { headers: basliklar, signal: AbortSignal.timeout(10_000) }).catch(() => null),
    ]);
    const j = (await r.json().catch(() => ({}))) as { display_phone_number?: string; verified_name?: string; quality_rating?: string; error?: { message?: string; code?: number } };
    if (!r.ok) return { ok: false, hata: `${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}${j.error?.code === 190 ? " → jeton geçersiz/süresi dolmuş" : ""}` };
    let uygulama: { id: string; ad: string } | undefined;
    if (ra && ra.ok) {
      const ja = (await ra.json().catch(() => ({}))) as { id?: string; name?: string };
      if (ja.id) uygulama = { id: ja.id, ad: ja.name || "" };
    }
    return { ok: true, numara: j.display_phone_number, ad: j.verified_name, kalite: j.quality_rating, uygulama };
  } catch (e) {
    return { ok: false, hata: (e as Error)?.message || String(e) };
  }
}

/**
 * Uygulamanın WhatsApp Business hesabına (WABA) webhook aboneliği. WHATSAPP_WABA_ID
 * tanımlıysa kontrol edilir; onar=true ile abone yapılır (gelen mesajlar ancak böyle düşer).
 */
export async function wabaAbonelik(onar = false): Promise<{ ok: boolean; wabaId?: string; abone?: boolean; alanlar?: string[]; uygulamalar?: { id: string; ad: string; alanlar: string[] }[]; hata?: string }> {
  const waba = (process.env.WHATSAPP_WABA_ID || "").trim();
  if (!waba) return { ok: false, hata: "WHATSAPP_WABA_ID tanımlı değil (Meta → WhatsApp → API Setup → WhatsApp Business Account ID)." };
  if (!token()) return { ok: false, hata: "WHATSAPP_TOKEN yok." };
  try {
    if (onar) {
      const r = await fetch(`${GRAPH}/${waba}/subscribed_apps`, { method: "POST", headers: { Authorization: `Bearer ${token()}` }, signal: AbortSignal.timeout(10_000) });
      const j = (await r.json().catch(() => ({}))) as { success?: boolean; error?: { message?: string; code?: number } };
      if (!r.ok || !j.success) return { ok: false, wabaId: waba, hata: `Abone olunamadı: ${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    }
    const r = await fetch(`${GRAPH}/${waba}/subscribed_apps`, { headers: { Authorization: `Bearer ${token()}` }, signal: AbortSignal.timeout(10_000) });
    const j = (await r.json().catch(() => ({}))) as { data?: { whatsapp_business_api_data?: { name?: string; id?: string }; subscribed_fields?: string[] }[]; error?: { message?: string; code?: number } };
    if (!r.ok) return { ok: false, wabaId: waba, hata: `${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    // Not: WhatsApp hesabı aboneliğinde Meta çoğu zaman alan listesi DÖNDÜRMEZ; boş liste "eksik" demek değildir.
    // "messages" alanı uygulama panelinde (WhatsApp → Configuration → Webhook fields) açılır.
    const alanlar = [...new Set((j.data || []).flatMap((d) => d.subscribed_fields || []))];
    const uygulamalar = (j.data || []).map((d) => ({ id: d.whatsapp_business_api_data?.id || "?", ad: d.whatsapp_business_api_data?.name || "?", alanlar: d.subscribed_fields || [] }));
    return { ok: true, wabaId: waba, abone: uygulamalar.length > 0, alanlar, uygulamalar };
  } catch (e) {
    return { ok: false, wabaId: waba, hata: (e as Error)?.message || String(e) };
  }
}

export interface WaGonderim { disId: string; yontem: "serbest" | "sablon"; sablon?: string }

async function sablonlaGonder(hesap: string, to: string, ad: string, metin: string): Promise<WaGonderim> {
  const hatalar: string[] = [];
  const adlar = await etkinSablonlar();
  if (adlar.length === 0) throw new Error("WhatsApp şablonu yok: Ayarlar → Mesajlar kartından \"Şablonu oluştur\" deyin (Meta onaylayınca kullanılır).");
  for (const sablon of adlar) {
    // Değişken sayısı Meta'dan biliniyorsa ona göre; bilinmiyorsa önce tek ({{1}} = metin), Meta
    // "değişken sayısı uyuşmuyor" derse ad + metin denenir.
    const bilgi = await sablonBilgisi(sablon).catch(() => null);
    const tek = [metin], cift = [ad || "Sayın Müşterimiz", metin];
    const denemeler = bilgi?.degisken === 2 ? [cift, tek] : [tek, cift];
    for (const params of denemeler) {
      const r = await grafGonder(hesap, {
        to, type: "template",
        template: { name: sablon, language: { code: sablonDili() }, components: [{ type: "body", parameters: params.map((t) => ({ type: "text", text: sablonParam(t) })) }] },
      });
      if (r.ok) return { disId: r.disId, yontem: "sablon", sablon };
      hatalar.push(`${sablon}: ${r.hata}`);
      if (r.kod !== 132000) break;
    }
  }
  throw new Error(`WhatsApp şablonla gönderilemedi — ${hatalar.join("; ")}`);
}

/** Mesaj gövdesini türüne göre metne çevirir; medya kimliğini döner. */
export function waGovde(m: WaMesaj): { metin: string; medya?: WaMedya & { tur: string } } {
  switch (m.type) {
    case "text": return { metin: m.text?.body || "" };
    case "image": return { metin: m.image?.caption || "", medya: { ...m.image, tur: "image" } };
    case "video": return { metin: m.video?.caption || "", medya: { ...m.video, tur: "video" } };
    case "audio": return { metin: "", medya: { ...m.audio, tur: "audio" } };
    case "document": return { metin: m.document?.caption || "", medya: { ...m.document, tur: "document" } };
    case "sticker": return { metin: "", medya: { ...m.sticker, tur: "sticker" } };
    case "location": {
      const l = m.location || {};
      const ad = [l.name, l.address].filter(Boolean).join(" — ");
      return { metin: `📍 Konum${ad ? ": " + ad : ""} https://maps.google.com/?q=${l.latitude},${l.longitude}` };
    }
    case "contacts":
      return {
        metin: (m.contacts || [])
          .map((c) => `👤 ${c.name?.formatted_name || "Kişi"} ${(c.phones || []).map((p) => p.phone || p.wa_id).filter(Boolean).join(", ")}`)
          .join("\n"),
      };
    case "reaction": return { metin: `Tepki: ${m.reaction?.emoji || ""}`.trim() };
    case "button": return { metin: m.button?.text || "" };
    case "interactive": return { metin: m.interactive?.button_reply?.title || m.interactive?.list_reply?.title || "" };
    case "unsupported": return { metin: "[Desteklenmeyen mesaj türü]" };
    default: return { metin: m.text?.body || (m.errors?.length ? `[Mesaj alınamadı: ${m.errors.map((e) => e.title || e.message).join(", ")}]` : "[Mesaj]") };
  }
}

async function medyaIndir(konusmaId: string, medya: WaMedya & { tur: string }): Promise<Ek> {
  const mime = medya.mime_type || "";
  const ad = medya.filename || `${medya.tur}-${(medya.id || "").slice(-8)}.${mimeUzanti(mime)}`;
  const ek: Ek = { tur: ekTuru(mime, medya.tur === "sticker" ? "sticker" : ad), ad, mime };
  if (!medya.id || !token()) return ek;
  try {
    const meta = await fetch(`${GRAPH}/${medya.id}`, { headers: { Authorization: `Bearer ${token()}` }, signal: AbortSignal.timeout(10_000) });
    if (!meta.ok) return ek;
    const j = (await meta.json()) as { url?: string; mime_type?: string; file_size?: number };
    if (!j.url) return ek;
    const indirilen = await ekIndir(j.url, { Authorization: `Bearer ${token()}` });
    if (!indirilen) return ek;
    const yol = await ekKaydet(konusmaId, ad, indirilen.mime || mime, indirilen.veri);
    if (yol) { ek.url = ekUrl(yol); ek.boyut = indirilen.veri.byteLength; ek.mime = indirilen.mime || mime; }
  } catch (e) {
    console.warn("WhatsApp medyası indirilemedi:", (e as Error)?.message);
  }
  return ek;
}

/** Webhook "messages" değerini işler; kaydedilen yeni mesaj sayısını döner. */
export async function gelenWhatsapp(v: WaValue): Promise<number> {
  const hesap = v.metadata?.phone_number_id || phoneId() || "whatsapp";
  let yeni = 0;
  for (const m of v.messages || []) {
    if (!m?.from || !m.id) continue;
    const adRehber = (v.contacts || []).find((c) => c.wa_id === m.from)?.profile?.name || "";
    const tel = normalizeWaNumber(m.from) || m.from;
    const k = await konusmaBulVeyaOlustur({ kanal: "whatsapp", hesap, disKimlik: tel, ad: adRehber || `+${tel}`, meta: { telefon: tel } });
    if (!k.musteriId) {
      const e = await musteriEsle({ telefon: tel }).catch(() => null);
      if (e) await konusmaGuncelle(k.id, { musteriId: e.musteriId, musteriTur: e.musteriTur });
    }
    const { metin, medya } = waGovde(m);
    const ekler: Ek[] = medya ? [await medyaIndir(k.id, medya)] : [];
    const at = m.timestamp && /^\d+$/.test(m.timestamp) ? new Date(Number(m.timestamp) * 1000) : new Date();
    const r = await mesajEkle({ konusmaId: k.id, yon: "gelen", govde: metin, ekler, disId: m.id, gonderen: adRehber || `+${tel}`, at });
    if (r.yeni) yeni++;
  }
  return yeni;
}

const DURUM: Record<string, string> = { sent: "gonderildi", delivered: "iletildi", read: "okundu", failed: "hata" };

/** Giden mesaj durumlarını (sent/delivered/read/failed) kaydeder; işlenen sayısını döner. */
export async function whatsappDurumlar(statuses: WaDurum[]): Promise<number> {
  let n = 0;
  for (const st of statuses || []) {
    if (!st?.id || !st.status || !DURUM[st.status]) continue;
    const hata = st.status === "failed"
      ? (st.errors || []).map((e) => `${e.title || e.message || "hata"}${e.error_data?.details ? ": " + e.error_data.details : ""} (kod ${e.code ?? "?"})`).join("; ") || "teslim edilemedi"
      : null;
    if (await mesajDurumDisId(st.id, DURUM[st.status], hata)) n++;
  }
  return n;
}

/**
 * Yanıt gönderir. Müşteri son 24 saatte yazdıysa serbest metin; yazmadıysa (ya da Meta 131047 ile
 * reddederse) WHATSAPP_TEMPLATE_SERBEST şablonuyla gider. Şablon yoksa pencere kapalıyken hata verir.
 */
export async function whatsappGonder(k: Konusma, metin: string): Promise<WaGonderim> {
  if (!whatsappConfigured()) throw new Error("WhatsApp Cloud API ayarlı değil (WHATSAPP_TOKEN / WHATSAPP_PHONE_ID).");
  const to = normalizeWaNumber(k.disKimlik) || k.disKimlik;
  const hesap = k.hesap || phoneId();
  if (pencereAcik(k)) {
    const r = await grafGonder(hesap, { to, type: "text", text: { preview_url: false, body: metin } });
    if (r.ok) return { disId: r.disId, yontem: "serbest" };
    if (r.kod !== 131047 || !(await whatsappSablonHazir())) throw new Error(`WhatsApp gönderilemedi: ${r.hata}`);
  }
  if (!(await whatsappSablonHazir())) {
    throw new Error("WhatsApp 24 saat penceresi kapalı: müşteri son 24 saatte yazmadığı için serbest mesaj gönderilemez. Müşteri tekrar yazınca yanıtlayabilirsiniz; ya da Ayarlar → Mesajlar kartından onaylı şablon oluşturun.");
  }
  const uzunluk = sablonParam(metin, Infinity).length;
  if (uzunluk > SABLON_METIN_AZAMI) {
    throw new Error(`Müşteri son 24 saatte yazmadığı için mesaj onaylı şablonla gidecek; şablon metni en fazla ${SABLON_METIN_AZAMI} karakter olabilir (şu an ${uzunluk}). Kısaltıp yeniden gönderin.`);
  }
  return sablonlaGonder(hesap, to, k.ad && !k.ad.startsWith("+") ? k.ad : "", metin);
}
