// Instagram DM bağlayıcısı (olga.cerceve profesyonel hesabı). İki yol:
//   • Instagram login (graph.instagram.com): INSTAGRAM_TOKEN (Instagram kullanıcı jetonu, 60 gün;
//     burada kendiliğinden tazelenir) + INSTAGRAM_ACCOUNT_ID + INSTAGRAM_APP_SECRET (webhook imzası).
//   • Facebook sayfası (graph.facebook.com): INSTAGRAM_TOKEN (sistem kullanıcısı jetonu, süresiz) +
//     INSTAGRAM_PAGE_ID + INSTAGRAM_ACCOUNT_ID; imza WhatsApp/META app secret ile.
// Webhook object="instagram", entry[].messaging[] olayları her iki yolda aynıdır.

import { createHash } from "node:crypto";
import { dbConfigured, konusmaBulVeyaOlustur, konusmaGuncelle, mesajBenzerVar, mesajEkle, senkOku, senkYaz } from "./db";
import { ekIndir, ekKaydet, ekTuru, ekUrl, mimeUzanti } from "./ek";
import { pencereAcik, type Ek, type Konusma, type Yon } from "./tur";

const FB = "https://graph.facebook.com/v20.0";
/**
 * Abone olunan webhook alanları. "standby" şart: müşteri Instagram uygulamasından (gelen kutusundan)
 * yanıtlanınca Meta konuşmanın kontrolünü Instagram gelen kutusuna verir (Handover Protocol) ve
 * sonraki mesajları "messaging" yerine "standby" altında gönderir; abone olunmazsa mesajlar hiç gelmez.
 */
export const IG_WEBHOOK_ALANLARI = "messages,messaging_postbacks,standby";
const IG = "https://graph.instagram.com/v21.0";
const IG_KOK = "https://graph.instagram.com";

const envJeton = () => (process.env.INSTAGRAM_TOKEN || process.env.INSTAGRAM_PAGE_TOKEN || "").trim();
const pageId = () => (process.env.INSTAGRAM_PAGE_ID || "").trim();
const accountId = () => (process.env.INSTAGRAM_ACCOUNT_ID || "").trim();
const parmakIzi = (s: string) => createHash("sha256").update(s).digest("hex").slice(0, 16);

export function instagramConfigured(): boolean {
  return Boolean(envJeton() && (pageId() || accountId()));
}
/** Sayfa kimliği yoksa Instagram login yolu (jeton 60 günlük, tazelenir). */
export function instagramLoginYolu(): boolean {
  return !pageId() && Boolean(accountId());
}

// ---- Jeton: tazelenmiş sürüm veri tabanında (mesaj_senk "ig:jeton"); env jetonu değişirse geçersiz sayılır.
interface SakliJeton { jeton: string; at: string; bitis: string | null; envIz: string }
let sakli: SakliJeton | null | undefined; // undefined: henüz okunmadı

async function sakliOku(): Promise<SakliJeton | null> {
  if (sakli !== undefined) return sakli;
  sakli = null;
  if (!dbConfigured()) return null;
  try {
    const r = await senkOku("ig:jeton");
    if (r?.deger) {
      const j = JSON.parse(r.deger) as SakliJeton;
      if (j?.jeton && j.envIz === parmakIzi(envJeton())) sakli = j;
    }
  } catch { /* env jetonuna düş */ }
  return sakli;
}
/** Testler için: jeton önbelleğini sıfırla. */
export function igJetonOnbellekSifirla() { sakli = undefined; }

/** Etkin jeton: aynı env jetonundan tazelenmiş sürüm varsa o, yoksa env. */
export async function jeton(): Promise<string> {
  const s = await sakliOku();
  return s?.jeton || envJeton();
}

/**
 * Instagram login jetonu 60 gün geçerlidir; 7 günde bir tazelenir (Meta 24 saatten yeni jetonu tazelemez).
 * Facebook sayfası yolunda jeton süresizdir, hiçbir şey yapılmaz.
 */
export async function igJetonTazele(zorla = false): Promise<{ ok: boolean; tazelendi?: boolean; bitis?: string | null; atlandi?: string; hata?: string }> {
  if (!instagramConfigured()) return { ok: false, hata: "Instagram ayarlı değil." };
  if (!instagramLoginYolu()) return { ok: true, tazelendi: false, atlandi: "Facebook sayfası yolunda jeton süresizdir." };
  if (!dbConfigured()) return { ok: false, hata: "Tazelenen jetonu saklamak için veri tabanı (DATABASE_URL) gerekli." };
  const s = await sakliOku();
  if (!zorla && s && Date.now() - new Date(s.at).getTime() < 7 * 86400_000) return { ok: true, tazelendi: false, bitis: s.bitis, atlandi: "7 gün dolmadı" };
  const mevcut = s?.jeton || envJeton();
  try {
    const r = await fetch(`${IG_KOK}/refresh_access_token?grant_type=ig_refresh_token&access_token=${encodeURIComponent(mevcut)}`, { signal: AbortSignal.timeout(10_000) });
    const j = (await r.json().catch(() => ({}))) as { access_token?: string; expires_in?: number; error?: { message?: string; code?: number } };
    if (!r.ok || !j.access_token) return { ok: false, bitis: s?.bitis, hata: `${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    const yeni: SakliJeton = { jeton: j.access_token, at: new Date().toISOString(), bitis: j.expires_in ? new Date(Date.now() + j.expires_in * 1000).toISOString() : null, envIz: parmakIzi(envJeton()) };
    await senkYaz("ig:jeton", JSON.stringify(yeni));
    sakli = yeni;
    return { ok: true, tazelendi: true, bitis: yeni.bitis };
  } catch (e) {
    return { ok: false, hata: (e as Error)?.message || String(e) };
  }
}

export async function igJetonDurumu(): Promise<{ yol: "instagram-login" | "facebook"; kaynak: "env" | "tazelenmis"; tazelendi: string | null; bitis: string | null }> {
  const s = await sakliOku();
  return { yol: instagramLoginYolu() ? "instagram-login" : "facebook", kaynak: s ? "tazelenmis" : "env", tazelendi: s?.at || null, bitis: s?.bitis || null };
}

// ---- Sayfa jetonu (Facebook yolu): Meta, mesaj gönderme ve profil sorgusunda sistem kullanıcısı jetonunu
// kabul etmez ("(#190) This method must be called with a Page Access Token"); sayfanın kendi jetonu
// /{page-id}?fields=access_token ile alınır (sistem kullanıcısından türeyen sayfa jetonu süresizdir).
let sayfaJetonuOnbellek: { jeton: string; iz: string } | null = null;
/** Testler için: sayfa jetonu önbelleğini sıfırla. */
export function igSayfaJetonuSifirla() { sayfaJetonuOnbellek = null; }

/** Facebook yolunda sayfa erişim jetonu (alınamazsa sistem jetonuna düşer); Instagram login yolunda etkin jeton. */
async function sayfaJetonu(): Promise<string> {
  const t = await jeton();
  if (!pageId() || !t) return t;
  const iz = parmakIzi(t);
  if (sayfaJetonuOnbellek?.iz === iz) return sayfaJetonuOnbellek.jeton;
  try {
    const r = await fetch(`${FB}/${pageId()}?fields=access_token&access_token=${encodeURIComponent(t)}`, { signal: AbortSignal.timeout(10_000) });
    const j = (await r.json().catch(() => ({}))) as { access_token?: string };
    if (r.ok && j.access_token) {
      sayfaJetonuOnbellek = { jeton: j.access_token, iz };
      return j.access_token;
    }
  } catch { /* sistem jetonuna düş */ }
  return t;
}

/** Gönderim / profil tabanı: Facebook sayfası üzerinden ya da doğrudan Instagram API. */
function taban(): { url: string; gonderim: string } {
  if (pageId()) return { url: FB, gonderim: `${FB}/${pageId()}/messages` };
  return { url: IG, gonderim: `${IG}/${accountId() || "me"}/messages` };
}

/** Jeton kontrolü (Ayarlar test kartı): sayfa/hesap adı. */
export async function instagramDurum(): Promise<{ ok: boolean; ad?: string; hata?: string; jeton?: Awaited<ReturnType<typeof igJetonDurumu>>; tazeleme?: Awaited<ReturnType<typeof igJetonTazele>> }> {
  if (!instagramConfigured()) return { ok: false, hata: "INSTAGRAM_TOKEN ile INSTAGRAM_PAGE_ID (ya da INSTAGRAM_ACCOUNT_ID) tanımlı değil." };
  try {
    // Sınama sırasında vadesi gelmişse jeton tazelenir (Instagram login yolu).
    const tazeleme = await igJetonTazele(false);
    const kim = pageId() || accountId();
    const alan = pageId() ? "name,instagram_business_account{username}" : "username";
    const r = await fetch(`${taban().url}/${kim}?fields=${encodeURIComponent(alan)}&access_token=${encodeURIComponent(await jeton())}`, { signal: AbortSignal.timeout(10_000) });
    const j = (await r.json().catch(() => ({}))) as { name?: string; username?: string; instagram_business_account?: { username?: string }; error?: { message?: string; code?: number } };
    if (!r.ok) return { ok: false, hata: `${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    const ig = j.instagram_business_account?.username || j.username;
    return { ok: true, ad: [j.name, ig ? "@" + ig : ""].filter(Boolean).join(" · "), jeton: await igJetonDurumu(), tazeleme };
  } catch (e) {
    return { ok: false, hata: (e as Error)?.message || String(e) };
  }
}

/**
 * Facebook Login yolunda Instagram mesaj webhook'ları ancak uygulama SAYFAYA abone olunca gelir
 * (POST /{page-id}/subscribed_apps, sayfa jetonuyla). Kontrol eder; onar=true ile abone yapar.
 */
export async function igSayfaAbonelik(onar = false): Promise<{ ok: boolean; abone?: boolean; uygulamalar?: { id: string; ad: string; alanlar: string[] }[]; hata?: string }> {
  if (!pageId() || !envJeton()) return { ok: false, hata: "INSTAGRAM_PAGE_ID ve INSTAGRAM_TOKEN gerekli (Instagram Login yolunda sayfa aboneliği yoktur)." };
  try {
    // Sayfa jetonu: sistem kullanıcısı / kullanıcı jetonuyla sayfanın kendi jetonu alınır (pages_manage_metadata).
    const sj = await sayfaJetonu();
    if (onar) {
      const r = await fetch(`${FB}/${pageId()}/subscribed_apps?subscribed_fields=${IG_WEBHOOK_ALANLARI}&access_token=${encodeURIComponent(sj)}`, { method: "POST", signal: AbortSignal.timeout(10_000) });
      const j = (await r.json().catch(() => ({}))) as { success?: boolean; error?: { message?: string; code?: number } };
      if (!r.ok || !j.success) return { ok: false, hata: `Sayfaya abone olunamadı: ${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    }
    const r = await fetch(`${FB}/${pageId()}/subscribed_apps?access_token=${encodeURIComponent(sj)}`, { signal: AbortSignal.timeout(10_000) });
    const j = (await r.json().catch(() => ({}))) as { data?: { id?: string; name?: string; subscribed_fields?: string[] }[]; error?: { message?: string; code?: number } };
    if (!r.ok) return { ok: false, hata: `${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    const uygulamalar = (j.data || []).map((d) => ({ id: d.id || "?", ad: d.name || "?", alanlar: d.subscribed_fields || [] }));
    return { ok: true, abone: uygulamalar.length > 0, uygulamalar };
  } catch (e) {
    return { ok: false, hata: (e as Error)?.message || String(e) };
  }
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
  let ad = "", kullanici = "", alindi = false;
  const t = await sayfaJetonu();
  if (t) {
    try {
      const r = await fetch(`${taban().url}/${igsid}?fields=name,username&access_token=${encodeURIComponent(t)}`, { signal: AbortSignal.timeout(8_000) });
      if (r.ok) {
        const j = (await r.json()) as { name?: string; username?: string };
        ad = j.name || ""; kullanici = j.username || ""; alindi = true;
      }
    } catch { /* profil alınamadı */ }
  }
  const v = { at: Date.now(), ad, kullanici };
  // Başarısız sorgu kısa süre (10 dk) önbelleklenir ki jeton/izin düzelince ad hemen gelsin.
  profilOnbellek.set(igsid, alindi ? v : { ...v, at: Date.now() - 6 * 3600_000 + 10 * 60_000 });
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
  // messaging: konuşma bizim uygulamada; standby: kontrol başka uygulamada (Instagram gelen kutusu) —
  // mesaj yine kaydedilir, yanıt gönderilirken kontrol geri alınır (take_thread_control).
  const olaylar = [
    ...(entry.messaging || []).map((ev) => ({ ev, standby: false })),
    ...(entry.standby || []).map((ev) => ({ ev, standby: true })),
  ];
  for (const { ev, standby } of olaylar) {
    const m = ev.message;
    if (!m?.mid || m.is_deleted) continue;
    const echo = Boolean(m.is_echo);
    // Echo: bizim hesaptan (Instagram uygulamasından ya da API'den) giden mesaj — karşı taraf alıcıdır.
    const karsi = echo ? ev.recipient?.id : ev.sender?.id;
    if (!karsi) continue;
    const profil = await igProfil(karsi);
    const ad = profil.ad || (profil.kullanici ? "@" + profil.kullanici : `Instagram ${karsi.slice(-6)}`);
    const k = await konusmaBulVeyaOlustur({ kanal: "instagram", hesap, disKimlik: karsi, ad, baslik: profil.kullanici ? "@" + profil.kullanici : "", meta: profil.kullanici ? { kullanici: profil.kullanici } : undefined });
    // Konuşma daha önce "Instagram 123456" gibi yer tutucu adla açıldıysa ve profil artık alınabiliyorsa adı düzelt.
    if ((profil.ad || profil.kullanici) && /^Instagram \d+$/.test(k.ad || "") && k.ad !== ad) {
      await konusmaGuncelle(k.id, { ad, meta: { ...k.meta, kullanici: profil.kullanici } });
      k.ad = ad;
    }
    // Kontrol durumu değiştiyse konuşmaya işle (standby → başka uygulamada; messaging → bizde).
    if (Boolean(k.meta.standby) !== standby) {
      const meta = { ...k.meta, standby, standbyAt: standby ? new Date().toISOString() : undefined };
      await konusmaGuncelle(k.id, { meta });
      k.meta = meta;
    }
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

// ---- Yedek yol: Meta Conversations API. Webhook gelmese de son Instagram konuşmaları çekilir,
// veri tabanında olmayan mesajlar işlenir (Gmail senkronu gibi: 45 sn'de en çok bir kez; zorla ile hemen).
const IG_SENK_ARALIK_MS = 45_000;
let igSenkSuruyor = false;
export interface IgSenkSonuc {
  hesap: string; yeni: number; konusma?: number; hata?: string; atlandi?: boolean;
  /** Teşhis için: API'de görünen son mesajlar (ad, zaman, yön, kısa metin). */
  son?: { ad: string; at: string; yon: Yon; ozet: string }[];
}
interface ApiKatilimci { id?: string; username?: string }
interface ApiEk { id?: string; name?: string; mime_type?: string; file_url?: string; image_data?: { url?: string; preview_url?: string }; video_data?: { url?: string } }
interface ApiMesaj { id?: string; created_time?: string; message?: string; from?: ApiKatilimci; attachments?: { data?: ApiEk[] } }
interface ApiKonusma { id?: string; updated_time?: string; participants?: { data?: ApiKatilimci[] }; messages?: { data?: ApiMesaj[] } }

function apiEkToIgEk(e: ApiEk): IgEk {
  if (e.image_data?.url) return { type: "image", payload: { url: e.image_data.url, title: e.name } };
  if (e.video_data?.url) return { type: "video", payload: { url: e.video_data.url, title: e.name } };
  return { type: (e.mime_type || "").startsWith("audio/") ? "audio" : "file", payload: { url: e.file_url, title: e.name } };
}

export async function instagramSenk(o: { zorla?: boolean } = {}): Promise<IgSenkSonuc> {
  const hesapAdi = "Instagram";
  if (!instagramConfigured() || !pageId()) return { hesap: hesapAdi, yeni: 0, atlandi: true };
  if (!dbConfigured()) return { hesap: hesapAdi, yeni: 0, hata: "DATABASE_URL tanımlı değil." };
  if (igSenkSuruyor) return { hesap: hesapAdi, yeni: 0, atlandi: true };
  igSenkSuruyor = true;
  try {
    const onceki = await senkOku("ig:senk");
    if (!o.zorla && onceki && Date.now() - new Date(onceki.at).getTime() < IG_SENK_ARALIK_MS) return { hesap: hesapAdi, yeni: 0, atlandi: true };
    // İlk çekimde son 7 gün; sonra son senkrondan 1 saat öncesine kadar (geç kalan webhook boşlukları kapanır).
    const esik = onceki ? new Date(onceki.at).getTime() - 3600_000 : Date.now() - 7 * 86400_000;
    const sj = await sayfaJetonu();
    const alanlar = "id,updated_time,participants,messages.limit(20){id,created_time,message,from,attachments}";
    const r = await fetch(`${FB}/${pageId()}/conversations?platform=instagram&fields=${encodeURIComponent(alanlar)}&limit=25&access_token=${encodeURIComponent(sj)}`, { signal: AbortSignal.timeout(20_000) });
    const j = (await r.json().catch(() => ({}))) as { data?: ApiKonusma[]; error?: { message?: string; code?: number } };
    if (!r.ok) return { hesap: hesapAdi, yeni: 0, hata: `${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    const bizim = new Set([accountId(), pageId()].filter(Boolean));
    let yeni = 0, konusmaSayisi = 0;
    const son: NonNullable<IgSenkSonuc["son"]> = [];
    for (const c of j.data || []) {
      if (c.updated_time && new Date(c.updated_time).getTime() < esik) continue;
      const karsi = (c.participants?.data || []).find((p) => p.id && !bizim.has(p.id));
      if (!karsi?.id) continue;
      konusmaSayisi++;
      const profil = await igProfil(karsi.id);
      const kullanici = profil.kullanici || karsi.username || "";
      const ad = profil.ad || (kullanici ? "@" + kullanici : `Instagram ${karsi.id.slice(-6)}`);
      const k = await konusmaBulVeyaOlustur({ kanal: "instagram", hesap: accountId() || pageId(), disKimlik: karsi.id, ad, baslik: kullanici ? "@" + kullanici : "", meta: { ...(kullanici ? { kullanici } : {}), apiKonusmaId: c.id } });
      if ((profil.ad || kullanici) && /^Instagram \d+$/.test(k.ad || "") && k.ad !== ad) { await konusmaGuncelle(k.id, { ad }); k.ad = ad; }
      const liste = [...(c.messages?.data || [])].sort((a, b) => new Date(a.created_time || 0).getTime() - new Date(b.created_time || 0).getTime());
      for (const m of liste) {
        if (!m.id || !m.from?.id) continue;
        const gelen = !bizim.has(m.from.id);
        const at = m.created_time ? new Date(m.created_time) : new Date();
        if (isNaN(at.getTime()) || at.getTime() < esik) continue;
        const govde = m.message || "";
        if (son.length < 3 && m === liste[liste.length - 1]) son.push({ ad: k.ad || karsi.id, at: at.toISOString(), yon: gelen ? "gelen" : "giden", ozet: govde.slice(0, 60) || "(ek)" });
        // Dış kimlik (mid) eşleşmese bile aynı metin ±3 dk içinde varsa tekrar yazma (webhook ile çift kayıt olmasın).
        if (await mesajBenzerVar(k.id, gelen ? "gelen" : "giden", govde, at)) continue;
        const ekler = await ekleriIndir(k.id, (m.attachments?.data || []).map(apiEkToIgEk));
        const e = await mesajEkle({
          konusmaId: k.id, yon: gelen ? "gelen" : "giden", govde, ekler, disId: m.id, at,
          gonderen: gelen ? (k.ad || karsi.id) : "Instagram uygulaması", durum: gelen ? "" : "gonderildi",
          sessiz: Date.now() - at.getTime() > 86400_000, // eski mesajlar okunmamış sayacını şişirmesin
        });
        if (e.yeni) yeni++;
      }
    }
    await senkYaz("ig:senk", JSON.stringify({ konusma: konusmaSayisi, yeni }));
    return { hesap: hesapAdi, yeni, konusma: konusmaSayisi, son };
  } catch (e) {
    return { hesap: hesapAdi, yeni: 0, hata: (e as Error)?.message || String(e) };
  } finally {
    igSenkSuruyor = false;
  }
}

// ---- Handover Protocol (Facebook sayfası yolu): konuşmanın kontrolü hangi uygulamada?
export interface ThreadSahibi { ok: boolean; appId?: string; ad?: string; bizde?: boolean; hata?: string }

/** Konuşmanın (IGSID) kontrolünü elinde tutan uygulama; bizim uygulama mı diye jetonun uygulamasıyla kıyaslar. */
export async function igThreadSahibi(igsid: string): Promise<ThreadSahibi> {
  if (!pageId()) return { ok: false, hata: "Instagram login yolunda thread sahibi sorgulanamaz." };
  try {
    const sj = await sayfaJetonu();
    const [r, ra] = await Promise.all([
      fetch(`${FB}/${pageId()}/thread_owner?recipient=${encodeURIComponent(igsid)}&access_token=${encodeURIComponent(sj)}`, { signal: AbortSignal.timeout(10_000) }),
      fetch(`${FB}/app?fields=id&access_token=${encodeURIComponent(await jeton())}`, { signal: AbortSignal.timeout(10_000) }).catch(() => null),
    ]);
    const j = (await r.json().catch(() => ({}))) as { data?: { thread_owner?: { app_id?: string } }[]; error?: { message?: string; code?: number } };
    if (!r.ok) return { ok: false, hata: `${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    const appId = j.data?.[0]?.thread_owner?.app_id || "";
    const bizimId = ra && ra.ok ? ((await ra.json().catch(() => ({}))) as { id?: string }).id || "" : "";
    let ad = "";
    if (appId) {
      const rn = await fetch(`${FB}/${appId}?fields=name&access_token=${encodeURIComponent(sj)}`, { signal: AbortSignal.timeout(8_000) }).catch(() => null);
      if (rn && rn.ok) ad = ((await rn.json().catch(() => ({}))) as { name?: string }).name || "";
    }
    return { ok: true, appId, ad, bizde: Boolean(appId && bizimId && appId === bizimId) };
  } catch (e) {
    return { ok: false, hata: (e as Error)?.message || String(e) };
  }
}

/** Konuşmanın kontrolünü bizim uygulamaya alır (uygulama sayfa ayarlarında "birincil alıcı" olmalı). */
export async function igThreadAl(igsid: string): Promise<{ ok: boolean; hata?: string }> {
  if (!pageId()) return { ok: false, hata: "Instagram login yolunda thread kontrolü yoktur." };
  try {
    const r = await fetch(`${FB}/${pageId()}/take_thread_control?access_token=${encodeURIComponent(await sayfaJetonu())}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ recipient: { id: igsid }, metadata: "olgasiparis" }),
      signal: AbortSignal.timeout(10_000),
    });
    const j = (await r.json().catch(() => ({}))) as { success?: boolean; error?: { message?: string; code?: number } };
    if (!r.ok || !j.success) return { ok: false, hata: `${j.error?.message || `HTTP ${r.status}`}${j.error?.code ? ` (kod ${j.error.code})` : ""}` };
    return { ok: true };
  } catch (e) {
    return { ok: false, hata: (e as Error)?.message || String(e) };
  }
}

const HANDOVER_HATA = /thread|handover|control|owner/i;

export async function instagramGonder(k: Konusma, metin: string): Promise<{ disId: string; kontrolAlindi?: boolean }> {
  if (!instagramConfigured()) throw new Error("Instagram mesajlaşma ayarlı değil (INSTAGRAM_TOKEN + INSTAGRAM_PAGE_ID).");
  if (!pencereAcik(k)) throw new Error("Instagram 24 saat penceresi kapalı: müşteri son 24 saatte yazmadığı için yanıt gönderilemez.");
  // Konuşma başka uygulamada (Instagram gelen kutusu) ise önce kontrolü al; alınamazsa yine de dene.
  let kontrolAlindi = false, kontrolDenendi = false, kontrolHatasi = "";
  const kontrolAl = async () => { kontrolDenendi = true; const r = await igThreadAl(k.disKimlik); kontrolAlindi = r.ok; kontrolHatasi = r.hata || ""; };
  if (k.meta?.standby && pageId()) await kontrolAl();
  const gonder = async () => {
    const r = await fetch(taban().gonderim, {
      method: "POST",
      headers: { Authorization: `Bearer ${await sayfaJetonu()}`, "Content-Type": "application/json" },
      body: JSON.stringify({ recipient: { id: k.disKimlik }, messaging_type: "RESPONSE", message: { text: metin } }),
      signal: AbortSignal.timeout(15_000),
    });
    const j = (await r.json().catch(() => ({}))) as { message_id?: string; recipient_id?: string; error?: { message?: string; code?: number; error_subcode?: number } };
    return { r, j };
  };
  let { r, j } = await gonder();
  if ((!r.ok || !j.message_id) && pageId() && !kontrolDenendi && HANDOVER_HATA.test(j.error?.message || "")) {
    // Kontrol başka uygulamadaymış: geri alıp bir kez daha dene.
    await kontrolAl();
    if (kontrolAlindi) ({ r, j } = await gonder());
  }
  if (!r.ok || !j.message_id) {
    if (j.error?.code === 190) igSayfaJetonuSifirla(); // jeton geçersiz/yanlış türde: bir sonraki denemede yeniden al
    const ipucu = kontrolHatasi ? ` — konuşma kontrolü alınamadı: ${kontrolHatasi} (Facebook sayfası ayarları → Gelişmiş mesajlaşma → Handover Protocol'de birincil alıcı bu uygulama olmalı)` : "";
    throw new Error(`Instagram gönderilemedi: ${j.error?.message || r.status}${j.error?.code ? ` (kod ${j.error.code})` : ""}${ipucu}`);
  }
  return { disId: j.message_id, kontrolAlindi };
}
