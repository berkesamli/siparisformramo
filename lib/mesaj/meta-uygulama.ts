// Meta uygulaması düzeyindeki webhook abonelikleri (GET/POST /{app-id}/subscriptions).
// Sayfa aboneliği (subscribed_apps) tek başına yetmez: uygulamanın kendisi de "instagram"
// nesnesinin "messages" alanına, bizim callback adresimizle abone olmalı. Bu ayar normalde
// Meta panelinden (Webhooks) yapılır; burada uygulama erişim jetonuyla (app-id|app-secret)
// sorgulanır ve gerekirse onarılır. Meta onarım sırasında callback adresini GET ile doğrular
// (hub.verify_token = INSTAGRAM_VERIFY_TOKEN ya da WHATSAPP_VERIFY_TOKEN).

import { IG_WEBHOOK_ALANLARI } from "./instagram";

const FB = "https://graph.facebook.com/v20.0";

export interface UygulamaAbonelik { nesne: string; url: string; aktif: boolean; alanlar: string[] }
export interface UygulamaWebhookDurumu {
  ok: boolean;
  uygulama?: { id: string; ad: string };
  abonelikler?: UygulamaAbonelik[];
  /** "instagram" nesnesi bizim adresimizle ve "messages" alanıyla abone mi? standby: Handover için gerekli alan da var mı? */
  instagram?: { abone: boolean; adresBizim: boolean; standby: boolean; alanlar: string[]; url?: string };
  hata?: string;
}

const jeton = () => (process.env.INSTAGRAM_TOKEN || process.env.INSTAGRAM_PAGE_TOKEN || process.env.WHATSAPP_TOKEN || "").trim();
const verifyToken = () => (process.env.INSTAGRAM_VERIFY_TOKEN || process.env.WHATSAPP_VERIFY_TOKEN || "").trim();
const secretAdaylari = () =>
  Array.from(new Set([process.env.META_APP_SECRET, process.env.WHATSAPP_APP_SECRET, process.env.INSTAGRAM_APP_SECRET].map((s) => (s || "").trim()).filter(Boolean)));

const hataMetni = (j: { error?: { message?: string; code?: number } } | undefined, r: Response) =>
  `${j?.error?.message || `HTTP ${r.status}`}${j?.error?.code ? ` (kod ${j.error.code})` : ""}`;

/** Jetonun ait olduğu uygulama. */
async function uygulama(): Promise<{ id: string; ad: string }> {
  const r = await fetch(`${FB}/app?fields=id,name&access_token=${encodeURIComponent(jeton())}`, { signal: AbortSignal.timeout(10_000) });
  const j = (await r.json().catch(() => ({}))) as { id?: string; name?: string; error?: { message?: string; code?: number } };
  if (!r.ok || !j.id) throw new Error(`Jetonun uygulaması bulunamadı: ${hataMetni(j, r)}`);
  return { id: j.id, ad: j.name || j.id };
}

interface AbonelikYaniti { data?: { object?: string; callback_url?: string; active?: boolean; fields?: { name?: string }[] }[]; error?: { message?: string; code?: number } }

/** Uygulama erişim jetonu (app-id|app-secret): env'deki secret adaylarından abonelik listesini okuyabilen ilki. */
async function uygulamaJetonu(appId: string): Promise<{ jeton: string; yanit: AbonelikYaniti }> {
  const adaylar = secretAdaylari();
  if (adaylar.length === 0) throw new Error("META_APP_SECRET / WHATSAPP_APP_SECRET tanımlı değil; uygulama webhook'u sorgulanamaz.");
  let sonHata = "";
  for (const s of adaylar) {
    const t = `${appId}|${s}`;
    const r = await fetch(`${FB}/${appId}/subscriptions?access_token=${encodeURIComponent(t)}`, { signal: AbortSignal.timeout(10_000) });
    const j = (await r.json().catch(() => ({}))) as AbonelikYaniti;
    if (r.ok) return { jeton: t, yanit: j };
    sonHata = hataMetni(j, r);
  }
  throw new Error(`Uygulama gizli anahtarı (App Secret) bu uygulamayla eşleşmiyor: ${sonHata}`);
}

function ozetle(yanit: AbonelikYaniti, callbackUrl: string): Pick<UygulamaWebhookDurumu, "abonelikler" | "instagram"> {
  const abonelikler = (yanit.data || []).map((d) => ({
    nesne: d.object || "?", url: d.callback_url || "", aktif: d.active !== false, alanlar: (d.fields || []).map((f) => f.name || "").filter(Boolean),
  }));
  const ig = abonelikler.find((a) => a.nesne === "instagram");
  const ayni = (a: string, b: string) => a.replace(/\/+$/, "").toLowerCase() === b.replace(/\/+$/, "").toLowerCase();
  return {
    abonelikler,
    instagram: ig
      ? { abone: ig.aktif && ig.alanlar.includes("messages"), adresBizim: ayni(ig.url, callbackUrl), standby: ig.alanlar.includes("standby"), alanlar: ig.alanlar, url: ig.url }
      : { abone: false, adresBizim: false, standby: false, alanlar: [] },
  };
}

/** Uygulamanın webhook aboneliklerini listeler (instagram / whatsapp_business_account / page). */
export async function uygulamaWebhookDurumu(callbackUrl: string): Promise<UygulamaWebhookDurumu> {
  if (!jeton()) return { ok: false, hata: "INSTAGRAM_TOKEN tanımlı değil." };
  try {
    const u = await uygulama();
    const { yanit } = await uygulamaJetonu(u.id);
    return { ok: true, uygulama: u, ...ozetle(yanit, callbackUrl) };
  } catch (e) {
    return { ok: false, hata: (e as Error)?.message || String(e) };
  }
}

/**
 * Uygulamayı "instagram" nesnesinin messages + messaging_postbacks alanlarına bizim adresimizle abone yapar.
 * Meta bu sırada callback adresini doğrular; verify token eşleşmezse Meta hata döner.
 */
export async function igUygulamaWebhookOnar(callbackUrl: string): Promise<UygulamaWebhookDurumu> {
  if (!jeton()) return { ok: false, hata: "INSTAGRAM_TOKEN tanımlı değil." };
  if (!verifyToken()) return { ok: false, hata: "INSTAGRAM_VERIFY_TOKEN (ya da WHATSAPP_VERIFY_TOKEN) tanımlı değil; Meta callback adresini doğrulayamaz." };
  try {
    const u = await uygulama();
    const { jeton: t } = await uygulamaJetonu(u.id);
    const govde = new URLSearchParams({
      object: "instagram",
      callback_url: callbackUrl,
      fields: IG_WEBHOOK_ALANLARI,
      verify_token: verifyToken(),
      include_values: "true",
      access_token: t,
    });
    const r = await fetch(`${FB}/${u.id}/subscriptions`, { method: "POST", body: govde, signal: AbortSignal.timeout(20_000) });
    const j = (await r.json().catch(() => ({}))) as { success?: boolean; error?: { message?: string; code?: number } };
    if (!r.ok || !j.success) return { ok: false, uygulama: u, hata: `Uygulama abone yapılamadı: ${hataMetni(j, r)}` };
    const { yanit } = await uygulamaJetonu(u.id);
    return { ok: true, uygulama: u, ...ozetle(yanit, callbackUrl) };
  } catch (e) {
    return { ok: false, hata: (e as Error)?.message || String(e) };
  }
}
