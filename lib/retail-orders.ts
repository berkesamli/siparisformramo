// Perakende (online çerçeve) siparişleri — Vercel Blob'da
// retail/<tarih>/<siparisNo>.json olarak tutulur. Yalnızca sunucu tarafında kullanılır.

import { findProfile } from "@/data/catalog";
import type { RetailStatus } from "@/data/perakende";
import { blobConfigured, istanbulDateKey } from "./orders";

// Perakende metre fiyatı = toptan liste fiyatı (USD/mt) × PERAKENDE_KATSAYI × kur.
// Katsayı gerekirse RETAIL_FACTOR env değişkeniyle değiştirilebilir.
const RETAIL_FACTOR = Number(process.env.RETAIL_FACTOR) || 1.5152;

export interface RetailFramePrice {
  found: boolean;
  code: string;
  resolvedCode?: string;
  tlPerM: number;
}

/** Profil koduna göre perakende çerçeve metre fiyatı (TL/m). */
export function retailFramePrice(code: string, usdRate: number): RetailFramePrice {
  const profile = findProfile(code);
  if (!profile || !(usdRate > 0)) {
    return { found: false, code, tlPerM: 0 };
  }
  const tlPerM =
    Math.round(profile.priceUSD * RETAIL_FACTOR * usdRate * 100) / 100;
  return { found: true, code, resolvedCode: profile.code, tlPerM };
}

export interface RetailItem {
  artWidth: number;
  artWidthUnit: "cm" | "mm";
  artHeight: number;
  artHeightUnit: "cm" | "mm";
  frameCode: string; // seri + renk kodu (veya OZEL)
  framePriceTL: number; // TL/m
  manualPrice: boolean;
  matType: string;
  matCode: string;
  matColor: string;
  matColorHex: string;
  doubleMat: boolean;
  innerMatType: string;
  innerMatColor: string;
  innerMatColorHex: string;
  altMontaj: string; // mm (çift paspartuda)
  zeminEnabled: boolean;
  zeminType: string;
  zeminColor: string;
  zeminColorHex: string;
  matTop: number;
  matRight: number;
  matBottom: number;
  matLeft: number;
  glassType: string;
  printType: string;
  frameCost: number;
  matCost: number;
  glassCost: number;
  printCost: number;
  itemTotal: number;
}

export interface SavedRetailOrder {
  orderId: string;
  dateKey: string; // YYYY-MM-DD (İstanbul)
  createdAt: string;
  updatedAt: string;
  status: RetailStatus;
  employee: string;
  customerName: string;
  customerPhone: string;
  customerEmail: string;
  usdRate: number;
  deliveryDate: string;
  notes: string;
  items: RetailItem[];
  gross: number;
  discount: number;
  total: number;
}

const orderPath = (dateKey: string, orderId: string) =>
  `retail/${dateKey}/${orderId}.json`;

export async function saveRetailOrder(order: SavedRetailOrder): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(orderPath(order.dateKey, order.orderId), JSON.stringify(order), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function getRetailOrder(
  dateKey: string,
  orderId: string
): Promise<SavedRetailOrder | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(orderPath(dateKey, orderId), {
      access: "private",
      useCache: false,
    });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as SavedRetailOrder;
  } catch {
    return null;
  }
}

export async function listRetailOrders(
  dateKeys: string[]
): Promise<SavedRetailOrder[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const orders: SavedRetailOrder[] = [];
  await Promise.all(
    dateKeys.map(async (key) => {
      try {
        const { blobs } = await list({ prefix: `retail/${key}/`, limit: 500 });
        await Promise.all(
          blobs.map(async (b) => {
            try {
              const r = await get(b.pathname, {
                access: "private",
                useCache: false,
              });
              if (!r || r.statusCode !== 200 || !r.stream) return;
              orders.push(
                JSON.parse(await new Response(r.stream).text()) as SavedRetailOrder
              );
            } catch {
              /* tek kayıt okunamazsa listeyi bozma */
            }
          })
        );
      } catch {
        /* gün klasörü yoksa geç */
      }
    })
  );
  return orders.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

/** Sıralı perakende sipariş numarası: PRK-2026-001. */
export async function nextRetailOrderId(): Promise<string> {
  const now = new Date();
  const year = istanbulDateKey(now).slice(0, 4);

  if (!blobConfigured()) {
    const pad = (n: number) => String(n).padStart(2, "0");
    const rand = Math.random().toString(36).slice(2, 5).toUpperCase();
    return `PRK-${year}-${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}-${rand}`;
  }

  const { get, put } = await import("@vercel/blob");
  const counterPath = `retail/counter-${year}.json`;
  let seq = 0;
  try {
    const r = await get(counterPath, { access: "private", useCache: false });
    if (r && r.statusCode === 200 && r.stream) {
      const data = JSON.parse(await new Response(r.stream).text());
      seq = Number(data.seq) || 0;
    }
  } catch {
    /* ilk sipariş — sayaç yok */
  }
  seq += 1;
  await put(counterPath, JSON.stringify({ seq }), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return `PRK-${year}-${String(seq).padStart(3, "0")}`;
}
