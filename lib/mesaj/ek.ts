// Mesaj ekleri (görsel, belge, ses) — Vercel Blob'da ÖZEL olarak saklanır ve
// yalnızca /api/mesaj/ek?p=… üzerinden (oturum + mesaj yetkisi) okunur.
// Depo yoksa ek kaydedilmez; mesaj yine kaydedilir, ek adı/açıklaması kalır.
//
// Yol: mesaj/ek/<konusmaId>/<rastgele>-<güvenli-ad>

import { blobConfigured } from "@/lib/orders";

export const EK_ON_EK = "mesaj/ek/";
export const EK_AZAMI_BAYT = 15 * 1024 * 1024;

const guvenliAd = (ad: string) =>
  String(ad || "dosya").normalize("NFKD").replace(/[^\w.\-]+/g, "_").replace(/_+/g, "_").slice(0, 80) || "dosya";

export function ekUrl(yol: string): string {
  return `/api/mesaj/ek?p=${encodeURIComponent(yol)}`;
}

export function ekYoluGecerli(yol: string): boolean {
  return typeof yol === "string" && yol.startsWith(EK_ON_EK) && !yol.includes("..") && yol.length < 400;
}

/** Baytları özel blob'a yazar; blob yoksa ya da hata olursa null döner (mesaj yine kaydedilir). */
export async function ekKaydet(konusmaId: string, ad: string, mime: string, veri: Uint8Array): Promise<string | null> {
  if (!blobConfigured()) return null;
  if (!veri || veri.byteLength === 0 || veri.byteLength > EK_AZAMI_BAYT) return null;
  try {
    const { put } = await import("@vercel/blob");
    const yol = `${EK_ON_EK}${konusmaId}/${Math.random().toString(36).slice(2, 8)}-${guvenliAd(ad)}`;
    await put(yol, Buffer.from(veri), {
      access: "private",
      contentType: mime || "application/octet-stream",
      addRandomSuffix: false,
      allowOverwrite: true,
    });
    return yol;
  } catch (e) {
    console.warn("Mesaj eki kaydedilemedi:", (e as Error)?.message);
    return null;
  }
}

/** Özel blob'dan okur (proxy rotası). */
export async function ekOku(yol: string): Promise<{ stream: ReadableStream; contentType: string; size?: number } | null> {
  if (!blobConfigured() || !ekYoluGecerli(yol)) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(yol, { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return { stream: r.stream, contentType: r.blob?.contentType || "application/octet-stream", size: r.blob?.size };
  } catch {
    return null;
  }
}

/** Dış bağlantıdan (Graph medya URL'si, IG CDN) baytları indirir; boyut sınırı aşılırsa null. */
export async function ekIndir(url: string, basliklar: Record<string, string> = {}, zamanAsimi = 20_000): Promise<{ veri: Uint8Array; mime: string } | null> {
  try {
    const r = await fetch(url, { headers: basliklar, signal: AbortSignal.timeout(zamanAsimi) });
    if (!r.ok) return null;
    const len = Number(r.headers.get("content-length") || 0);
    if (len > EK_AZAMI_BAYT) return null;
    const buf = new Uint8Array(await r.arrayBuffer());
    if (buf.byteLength > EK_AZAMI_BAYT) return null;
    return { veri: buf, mime: (r.headers.get("content-type") || "application/octet-stream").split(";")[0].trim() };
  } catch {
    return null;
  }
}

export function ekTuru(mime: string, ad = ""): "image" | "document" | "audio" | "video" | "sticker" | "other" {
  const m = (mime || "").toLowerCase();
  if (m.startsWith("image/webp") && /sticker/i.test(ad)) return "sticker";
  if (m.startsWith("image/")) return "image";
  if (m.startsWith("audio/")) return "audio";
  if (m.startsWith("video/")) return "video";
  if (m === "application/pdf" || /word|excel|sheet|text\/|zip|octet/.test(m)) return "document";
  return "other";
}

export function mimeUzanti(mime: string): string {
  const m = (mime || "").toLowerCase().split(";")[0];
  const tablo: Record<string, string> = {
    "image/jpeg": "jpg", "image/png": "png", "image/webp": "webp", "image/gif": "gif",
    "application/pdf": "pdf", "audio/ogg": "ogg", "audio/mpeg": "mp3", "audio/mp4": "m4a",
    "video/mp4": "mp4", "text/plain": "txt",
  };
  return tablo[m] || (m.split("/")[1] || "bin").replace(/[^a-z0-9]/g, "").slice(0, 5) || "bin";
}
