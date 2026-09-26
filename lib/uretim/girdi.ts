// API gövdelerinden gelen iş / kalem / not alanlarını temizler (sunucu tarafı).
// İstemciden gelen hiçbir alan doğrudan veri tabanına yazılmaz.

import type { RetailItem } from "@/lib/retail-orders";
import { DURUM_LABELS, SAAT_RE, TARIH_RE, type IsDurum, type IsKalem, type Sube } from "./tur";

export const s = (v: unknown, max = 200) => String(v ?? "").slice(0, max);
export const r2 = (n: unknown) => Math.round((Number(n) || 0) * 100) / 100;

export function subeOku(v: unknown): Sube | undefined {
  if (v === undefined) return undefined;
  const t = String(v || "");
  return t === "ankara" || t === "istanbul" ? t : "";
}

export function tarihOku(v: unknown): string | null | undefined {
  if (v === undefined) return undefined;
  if (v === null || v === "") return null;
  const t = String(v).slice(0, 10);
  return TARIH_RE.test(t) ? t : undefined;
}

export function saatOku(v: unknown): string | undefined {
  if (v === undefined) return undefined;
  const t = String(v || "").slice(0, 5);
  return t === "" || SAAT_RE.test(t) ? t : undefined;
}

export function durumOku(v: unknown): IsDurum | undefined {
  if (v === undefined) return undefined;
  const t = String(v || "");
  return Object.prototype.hasOwnProperty.call(DURUM_LABELS, t) ? (t as IsDurum) : undefined;
}

function retailOku(raw: any): RetailItem | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const it: RetailItem = {
    artWidth: Number(raw.artWidth) || 0,
    artWidthUnit: raw.artWidthUnit === "mm" ? "mm" : "cm",
    artHeight: Number(raw.artHeight) || 0,
    artHeightUnit: raw.artHeightUnit === "mm" ? "mm" : "cm",
    frameCode: s(raw.frameCode, 60),
    framePriceTL: r2(raw.framePriceTL),
    manualPrice: Boolean(raw.manualPrice),
    matType: s(raw.matType, 60) || "Paspartu Yok",
    matCode: s(raw.matCode, 20),
    matColor: s(raw.matColor, 20) || "-",
    matColorHex: s(raw.matColorHex, 20) || "-",
    doubleMat: Boolean(raw.doubleMat),
    innerMatType: s(raw.innerMatType, 60) || "-",
    innerMatColor: s(raw.innerMatColor, 20) || "-",
    innerMatColorHex: s(raw.innerMatColorHex, 20) || "-",
    altMontaj: s(raw.altMontaj, 10) || "-",
    zeminEnabled: Boolean(raw.zeminEnabled),
    zeminType: s(raw.zeminType, 60) || "-",
    zeminColor: s(raw.zeminColor, 20) || "-",
    zeminColorHex: s(raw.zeminColorHex, 20) || "-",
    matTop: Number(raw.matTop) || 0,
    matRight: Number(raw.matRight) || 0,
    matBottom: Number(raw.matBottom) || 0,
    matLeft: Number(raw.matLeft) || 0,
    pencereSayisi: Math.min(9, Math.max(1, Math.round(Number(raw.pencereSayisi) || 1))),
    pencereDuzen: s(raw.pencereDuzen, 8),
    pencereAralik: r2(raw.pencereAralik),
    kasa: Boolean(raw.kasa),
    glassType: s(raw.glassType, 60) || "Cam Yok",
    printType: s(raw.printType, 60) || "Baskı Yok",
    frameCost: r2(raw.frameCost),
    matCost: r2(raw.matCost),
    glassCost: r2(raw.glassCost),
    printCost: r2(raw.printCost),
    itemTotal: r2(raw.itemTotal),
  };
  return it.artWidth > 0 && it.artHeight > 0 ? it : undefined;
}

/** Kalem listesi: sku + adet zorunlu; ölçü/paspartu/cam serbest metin; retail (föy verisi) isteğe bağlı. */
export function kalemleriOku(raw: unknown): IsKalem[] {
  if (!Array.isArray(raw)) return [];
  return raw
    .slice(0, 40)
    .map((k: any): IsKalem => {
      const mm = (v: any) => (v && Number(v.w) > 0 && Number(v.h) > 0 ? { w: Number(v.w), h: Number(v.h) } : undefined);
      const out: IsKalem = {
        sku: s(k?.sku, 60).trim(),
        adet: Math.max(1, Math.min(999, Math.round(Number(k?.adet) || 1))),
        ozet: s(k?.ozet, 300).trim(),
      };
      const eser = mm(k?.eserMm); if (eser) out.eserMm = eser;
      const dis = mm(k?.disMm); if (dis) out.disMm = dis;
      if (k?.icerik) out.icerik = s(k.icerik, 80);
      if (k?.yon) out.yon = s(k.yon, 20);
      if (k?.paspartu) out.paspartu = s(k.paspartu, 120);
      if (k?.cam) out.cam = s(k.cam, 60);
      if (k?.pay !== undefined) out.pay = Boolean(k.pay);
      if (k?.fiyat !== undefined) out.fiyat = r2(k.fiyat);
      const retail = retailOku(k?.retail); if (retail) out.retail = retail;
      return out;
    })
    .filter((k) => k.sku || k.ozet);
}
