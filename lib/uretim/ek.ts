// Üretim işi ekleri (föy PDF'i, müşteri görseli, dış föy) — Vercel Blob'da
// ÖZEL saklanır; yalnızca /api/uretim/ek?p=… (oturum + üretim yetkisi) okur.
//
// Yollar: uretim/ek/<isId>/<rastgele>-<güvenli-ad>   (yüklenen ekler)
//         uretim/foy/<isId>.pdf                        (üretilen föy)

import { blobConfigured } from "@/lib/orders";

export const EK_ON_EK = "uretim/";
export const EK_AZAMI_BAYT = 15 * 1024 * 1024;

const guvenliAd = (ad: string) =>
  String(ad || "dosya").normalize("NFKD").replace(/[^\w.\-]+/g, "_").replace(/_+/g, "_").slice(0, 80) || "dosya";

export function ekUrl(yol: string, indir = false): string {
  return `/api/uretim/ek?p=${encodeURIComponent(yol)}${indir ? "&indir=1" : ""}`;
}

export function ekYoluGecerli(yol: string): boolean {
  return typeof yol === "string" && yol.startsWith(EK_ON_EK) && !yol.includes("..") && yol.length < 400;
}

export function ekTuru(mime: string): "pdf" | "image" | "file" {
  const m = (mime || "").toLowerCase();
  if (m === "application/pdf") return "pdf";
  if (m.startsWith("image/")) return "image";
  return "file";
}

/** Yüklenen eki yazar; blob yoksa / hata olursa null. */
export async function ekKaydet(isId: string, ad: string, mime: string, veri: Uint8Array): Promise<string | null> {
  if (!blobConfigured()) return null;
  if (!veri || veri.byteLength === 0 || veri.byteLength > EK_AZAMI_BAYT) return null;
  try {
    const { put } = await import("@vercel/blob");
    const yol = `${EK_ON_EK}ek/${isId}/${Math.random().toString(36).slice(2, 8)}-${guvenliAd(ad)}`;
    await put(yol, Buffer.from(veri), { access: "private", contentType: mime || "application/octet-stream", addRandomSuffix: false, allowOverwrite: true });
    return yol;
  } catch (e) {
    console.warn("Üretim eki kaydedilemedi:", (e as Error)?.message);
    return null;
  }
}

/** Üretilen föy PDF'ini sabit yola yazar (yeniden üretimde üzerine yazar). */
export async function foyKaydet(isId: string, pdf: Buffer): Promise<string | null> {
  if (!blobConfigured()) return null;
  try {
    const { put } = await import("@vercel/blob");
    const yol = `${EK_ON_EK}foy/${isId}.pdf`;
    await put(yol, pdf, { access: "private", contentType: "application/pdf", addRandomSuffix: false, allowOverwrite: true });
    return yol;
  } catch (e) {
    console.warn("Föy kaydedilemedi:", (e as Error)?.message);
    return null;
  }
}

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

export async function ekSil(yol: string): Promise<void> {
  if (!blobConfigured() || !ekYoluGecerli(yol)) return;
  try {
    const { del } = await import("@vercel/blob");
    await del(yol);
  } catch {
    /* yoksa sorun değil */
  }
}
