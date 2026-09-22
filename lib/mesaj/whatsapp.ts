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

/** Serbest metin yanıtı. Müşteri 24 saattir yazmadıysa Meta reddeder; önce pencere kontrol edilir. */
export async function whatsappGonder(k: Konusma, metin: string): Promise<{ disId: string }> {
  if (!whatsappConfigured()) throw new Error("WhatsApp Cloud API ayarlı değil (WHATSAPP_TOKEN / WHATSAPP_PHONE_ID).");
  if (!pencereAcik(k)) throw new Error("WhatsApp 24 saat penceresi kapalı: müşteri son 24 saatte yazmadığı için serbest mesaj gönderilemez. Müşteri tekrar yazınca yanıtlayabilirsiniz.");
  const to = normalizeWaNumber(k.disKimlik) || k.disKimlik;
  const r = await fetch(`${GRAPH}/${k.hesap || phoneId()}/messages`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token()}`, "Content-Type": "application/json" },
    body: JSON.stringify({ messaging_product: "whatsapp", recipient_type: "individual", to, type: "text", text: { preview_url: false, body: metin } }),
    signal: AbortSignal.timeout(15_000),
  });
  const j = (await r.json().catch(() => ({}))) as { messages?: { id: string }[]; error?: { message?: string; code?: number; error_data?: { details?: string } } };
  if (!r.ok || !j.messages?.[0]?.id) {
    const e = j.error;
    throw new Error(`WhatsApp gönderilemedi: ${e?.message || r.status}${e?.error_data?.details ? " — " + e.error_data.details : ""}`);
  }
  return { disId: j.messages[0].id };
}
