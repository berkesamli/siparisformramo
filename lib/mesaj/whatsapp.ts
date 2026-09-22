// WhatsApp Cloud API bağlayıcısı (gelen kutusu).
//   gelenWhatsapp(value)     — webhook "messages" alanını konuşma + mesaja çevirir
//   whatsappDurumlar(statuses) — giden mesajların teslim/okundu durumunu işler
//   whatsappGonder(konusma, metin) — serbest metin yanıtı (24 saat penceresi içinde)
//
// Aynı Meta uygulaması ve WHATSAPP_TOKEN / WHATSAPP_PHONE_ID kullanılır.
// Medya: Graph'tan indirilip özel Blob'a yazılır (bkz. ek.ts).

import { normalizeWaNumber, whatsappConfigured } from "@/lib/whatsapp-pdf";
import { konusmaBulVeyaOlustur, konusmaGuncelle, mesajDurumDisId, mesajEkle } from "./db";
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
export function whatsappSablonHazir(): boolean {
  return whatsappConfigured() && serbestSablonAdlari().length > 0;
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
export async function wabaAbonelik(onar = false): Promise<{ ok: boolean; wabaId?: string; abone?: boolean; alanlar?: string[]; hata?: string }> {
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
    const alanlar = [...new Set((j.data || []).flatMap((d) => d.subscribed_fields || []))];
    return { ok: true, wabaId: waba, abone: (j.data || []).length > 0, alanlar };
  } catch (e) {
    return { ok: false, wabaId: waba, hata: (e as Error)?.message || String(e) };
  }
}

export interface WaGonderim { disId: string; yontem: "serbest" | "sablon"; sablon?: string }

async function sablonlaGonder(hesap: string, to: string, ad: string, metin: string): Promise<WaGonderim> {
  const hatalar: string[] = [];
  for (const sablon of serbestSablonAdlari()) {
    // Önce tek değişken ({{1}} = metin); Meta "değişken sayısı uyuşmuyor" derse ad + metin denenir.
    for (const params of [[metin], [ad || "Sayın Müşterimiz", metin]]) {
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
    if (r.kod !== 131047 || !whatsappSablonHazir()) throw new Error(`WhatsApp gönderilemedi: ${r.hata}`);
  }
  if (!whatsappSablonHazir()) {
    throw new Error("WhatsApp 24 saat penceresi kapalı: müşteri son 24 saatte yazmadığı için serbest mesaj gönderilemez. Müşteri tekrar yazınca yanıtlayabilirsiniz; ya da Meta onaylı bir şablon tanımlayın (WHATSAPP_TEMPLATE_SERBEST).");
  }
  const uzunluk = sablonParam(metin, Infinity).length;
  if (uzunluk > SABLON_METIN_AZAMI) {
    throw new Error(`Müşteri son 24 saatte yazmadığı için mesaj onaylı şablonla gidecek; şablon metni en fazla ${SABLON_METIN_AZAMI} karakter olabilir (şu an ${uzunluk}). Kısaltıp yeniden gönderin.`);
  }
  return sablonlaGonder(hesap, to, k.ad && !k.ad.startsWith("+") ? k.ad : "", metin);
}
