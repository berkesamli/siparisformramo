// Instagram DM bağlayıcısı (olga.cerceve profesyonel hesabı).
// Meta "Messenger Platform for Instagram": webhook object="instagram",
// entry[].messaging[] olayları; gönderim POST /{page-id}/messages.
//
// Ortam değişkenleri:
//   INSTAGRAM_TOKEN        — sayfa erişim jetonu (instagram_manage_messages izniyle)
//   INSTAGRAM_PAGE_ID      — Instagram hesabının bağlı olduğu Facebook sayfası
//   INSTAGRAM_ACCOUNT_ID   — Instagram işletme hesabı kimliği (webhook entry.id)
// Sayfa kimliği yoksa "Instagram Login" API'si (graph.instagram.com) kullanılır.

import { konusmaBulVeyaOlustur, konusmaGuncelle, mesajEkle } from "./db";
import { ekIndir, ekKaydet, ekTuru, ekUrl, mimeUzanti } from "./ek";
import { pencereAcik, type Ek, type Konusma } from "./tur";

const FB = "https://graph.facebook.com/v20.0";
const IG = "https://graph.instagram.com/v21.0";

const token = () => (process.env.INSTAGRAM_TOKEN || process.env.INSTAGRAM_PAGE_TOKEN || "").trim();
const pageId = () => (process.env.INSTAGRAM_PAGE_ID || "").trim();
const accountId = () => (process.env.INSTAGRAM_ACCOUNT_ID || "").trim();

export function instagramConfigured(): boolean {
  return Boolean(token() && (pageId() || accountId()));
}

/** Gönderim / profil tabanı: Facebook sayfası üzerinden ya da doğrudan Instagram API. */
function taban(): { url: string; gonderim: string } {
  if (pageId()) return { url: FB, gonderim: `${FB}/${pageId()}/messages` };
  return { url: IG, gonderim: `${IG}/${accountId() || "me"}/messages` };
}

export interface IgEk { type?: string; payload?: { url?: string; title?: string; sticker_id?: number } }
export interface IgOlay {
  sender?: { id?: string };
  recipient?: { id?: string };
  timestamp?: number;
  message?: {
    mid?: string;
    text?: string;
    attachments?: IgEk[];
    is_echo?: boolean;
    is_deleted?: boolean;
    is_unsupported?: boolean;
    reply_to?: { mid?: string; story?: { url?: string; id?: string } };
  };
  read?: { mid?: string };
  reaction?: { mid?: string; action?: string; emoji?: string; reaction?: string };
  postback?: { title?: string; payload?: string; mid?: string };
}
export interface IgEntry { id?: string; time?: number; messaging?: IgOlay[]; standby?: IgOlay[] }

const profilOnbellek = new Map<string, { at: number; ad: string; kullanici: string }>();

/** IGSID → ad / kullanıcı adı (izin yoksa boş döner; hata sessiz geçilir). */
export async function igProfil(igsid: string): Promise<{ ad: string; kullanici: string }> {
  const c = profilOnbellek.get(igsid);
  if (c && Date.now() - c.at < 6 * 3600_000) return c;
  let ad = "", kullanici = "";
  if (token()) {
    try {
      const r = await fetch(`${taban().url}/${igsid}?fields=name,username&access_token=${encodeURIComponent(token())}`, { signal: AbortSignal.timeout(8_000) });
      if (r.ok) {
        const j = (await r.json()) as { name?: string; username?: string };
        ad = j.name || ""; kullanici = j.username || "";
      }
    } catch { /* profil alınamadı */ }
  }
  const v = { at: Date.now(), ad, kullanici };
  profilOnbellek.set(igsid, v);
  return v;
}

function ekAdi(e: IgEk, i: number): string {
  const t = e.type || "file";
  const u = e.payload?.url || "";
  const uz = (u.split("?")[0].split(".").pop() || "").toLowerCase();
  return e.payload?.title || `${t}-${i + 1}${uz && uz.length <= 4 ? "." + uz : ""}`;
}

async function ekleriIndir(konusmaId: string, ekler: IgEk[]): Promise<Ek[]> {
  const out: Ek[] = [];
  for (const [i, e] of (ekler || []).slice(0, 6).entries()) {
    const t = e.type || "other";
    const ad = ekAdi(e, i);
    const ek: Ek = { tur: t === "image" ? "image" : t === "video" || t === "reel" || t === "ig_reel" ? "video" : t === "audio" ? "audio" : "other", ad, url: e.payload?.url };
    if (t === "share" || t === "story_mention" || t === "reel" || t === "ig_reel") { ek.ad = t === "share" ? "Paylaşılan gönderi" : t === "story_mention" ? "Hikâyede bahsetme" : "Reels"; out.push(ek); continue; }
    if (e.payload?.url) {
      const ind = await ekIndir(e.payload.url);
      if (ind) {
        const yol = await ekKaydet(konusmaId, ad.includes(".") ? ad : `${ad}.${mimeUzanti(ind.mime)}`, ind.mime, ind.veri);
        if (yol) { ek.url = ekUrl(yol); ek.mime = ind.mime; ek.boyut = ind.veri.byteLength; ek.tur = ekTuru(ind.mime, ad); }
      }
    }
    out.push(ek);
  }
  return out;
}

/** Bir webhook girdisini işler; kaydedilen yeni mesaj sayısını döner. */
export async function gelenInstagram(entry: IgEntry): Promise<number> {
  const hesap = entry.id || accountId() || "instagram";
  let yeni = 0;
  for (const ev of entry.messaging || []) {
    const m = ev.message;
    if (!m?.mid || m.is_deleted) continue;
    const echo = Boolean(m.is_echo);
    // Echo: bizim hesaptan (Instagram uygulamasından ya da API'den) giden mesaj — karşı taraf alıcıdır.
    const karsi = echo ? ev.recipient?.id : ev.sender?.id;
    if (!karsi) continue;
    const profil = await igProfil(karsi);
    const ad = profil.ad || (profil.kullanici ? "@" + profil.kullanici : `Instagram ${karsi.slice(-6)}`);
    const k = await konusmaBulVeyaOlustur({ kanal: "instagram", hesap, disKimlik: karsi, ad, baslik: profil.kullanici ? "@" + profil.kullanici : "", meta: profil.kullanici ? { kullanici: profil.kullanici } : undefined });
    let metin = m.text || "";
    if (m.is_unsupported && !metin) metin = "[Desteklenmeyen mesaj türü]";
    if (m.reply_to?.story?.url) metin = `(Hikâyeye yanıt) ${metin}`.trim();
    const ekler = await ekleriIndir(k.id, m.attachments || []);
    const at = ev.timestamp ? new Date(ev.timestamp > 1e12 ? ev.timestamp : ev.timestamp * 1000) : new Date();
    const r = await mesajEkle({
      konusmaId: k.id, yon: echo ? "giden" : "gelen", govde: metin, ekler, disId: m.mid,
      gonderen: echo ? "Instagram uygulaması" : ad, durum: echo ? "gonderildi" : "", at,
    });
    if (r.yeni) yeni++;
  }
  return yeni;
}

/** Konuşmadaki profil bilgisini tazeler (elle "yenile"). */
export async function igProfilTazele(k: Konusma): Promise<void> {
  const p = await igProfil(k.disKimlik);
  if (p.ad || p.kullanici) await konusmaGuncelle(k.id, { ad: p.ad || k.ad, meta: { ...k.meta, kullanici: p.kullanici } });
}

export async function instagramGonder(k: Konusma, metin: string): Promise<{ disId: string }> {
  if (!instagramConfigured()) throw new Error("Instagram mesajlaşma ayarlı değil (INSTAGRAM_TOKEN + INSTAGRAM_PAGE_ID).");
  if (!pencereAcik(k)) throw new Error("Instagram 24 saat penceresi kapalı: müşteri son 24 saatte yazmadığı için yanıt gönderilemez.");
  const r = await fetch(taban().gonderim, {
    method: "POST",
    headers: { Authorization: `Bearer ${token()}`, "Content-Type": "application/json" },
    body: JSON.stringify({ recipient: { id: k.disKimlik }, messaging_type: "RESPONSE", message: { text: metin } }),
    signal: AbortSignal.timeout(15_000),
  });
  const j = (await r.json().catch(() => ({}))) as { message_id?: string; recipient_id?: string; error?: { message?: string; code?: number; error_subcode?: number } };
  if (!r.ok || !j.message_id) {
    throw new Error(`Instagram gönderilemedi: ${j.error?.message || r.status}${j.error?.code ? ` (kod ${j.error.code})` : ""}`);
  }
  return { disId: j.message_id };
}
