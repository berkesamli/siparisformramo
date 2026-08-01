// Sipariş kayıtları — Vercel Blob'da orders/<tarih>/<siparisNo>.json olarak tutulur.
// Yalnızca sunucu tarafında kullanılır.

import type { OrderLine } from "./notify";

export type OrderStatus = "olusturuldu" | "hazirlaniyor" | "tamamlandi";

export const STATUS_LABELS: Record<OrderStatus, string> = {
  olusturuldu: "Oluşturuldu",
  hazirlaniyor: "Hazırlanıyor",
  tamamlandi: "Tamamlandı",
};

export interface SavedOrder {
  orderId: string;
  dateKey: string; // YYYY-MM-DD (İstanbul)
  createdAt: string; // ISO
  updatedAt: string; // ISO
  status: OrderStatus;
  employee: string;
  customer: string;
  note: string;
  rate: number;
  euroRate: number;
  discountPct: number;
  vatApplied: boolean;
  lines: OrderLine[];
  gross: number;
  discount: number;
  vatAmount: number;
  net: number;
  rows?: unknown[]; // formun ham satırları — düzenleme için
}

export function sanitizeLines(raw: unknown): OrderLine[] {
  if (!Array.isArray(raw)) return [];
  return raw
    .map((l: any) => ({
      name: String(l?.name || "").slice(0, 200),
      unitText: String(l?.unitText || "").slice(0, 200),
      unitPriceTL: Number(l?.unitPriceTL) || 0,
      lineTotal: Number(l?.lineTotal) || 0,
    }))
    .filter((l) => l.name);
}

export function computeTotals(lines: OrderLine[], discountPct: number, vatApplied: boolean) {
  const r2 = (n: number) => Math.round(n * 100) / 100;
  const gross = r2(lines.reduce((s, l) => s + l.lineTotal, 0));
  const discount = r2(gross * (Math.max(0, discountPct) / 100));
  const afterDiscount = r2(Math.max(0, gross - discount));
  const vatAmount = vatApplied ? r2(afterDiscount * 0.2) : 0;
  const net = r2(afterDiscount + vatAmount);
  return { gross, discount, vatAmount, net };
}

export function blobConfigured(): boolean {
  return Boolean(process.env.BLOB_STORE_ID || process.env.BLOB_READ_WRITE_TOKEN);
}

export function istanbulDateKey(d = new Date()): string {
  return d.toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" });
}

function orderPath(dateKey: string, orderId: string): string {
  return `orders/${dateKey}/${orderId}.json`;
}

export async function saveOrder(order: SavedOrder): Promise<boolean> {
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

export async function getOrder(
  dateKey: string,
  orderId: string
): Promise<SavedOrder | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const result = await get(orderPath(dateKey, orderId), {
      access: "private",
      useCache: false,
    });
    if (!result || result.statusCode !== 200 || !result.stream) return null;
    const text = await new Response(result.stream).text();
    return JSON.parse(text) as SavedOrder;
  } catch {
    return null;
  }
}

/** Verilen tarihlerdeki (YYYY-MM-DD) tüm siparişleri getirir, yeniden eskiye sıralar. */
export async function listOrders(dateKeys: string[]): Promise<SavedOrder[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const orders: SavedOrder[] = [];
  await Promise.all(
    dateKeys.map(async (key) => {
      try {
        const { blobs } = await list({ prefix: `orders/${key}/`, limit: 500 });
        await Promise.all(
          blobs.map(async (b) => {
            try {
              const result = await get(b.pathname, {
                access: "private",
                useCache: false,
              });
              if (!result || result.statusCode !== 200 || !result.stream) return;
              const text = await new Response(result.stream).text();
              orders.push(JSON.parse(text) as SavedOrder);
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

/** Bugünden geriye n günlük tarih anahtarları (İstanbul saati). */
export function lastNDateKeys(n: number): string[] {
  const keys: string[] = [];
  const now = new Date();
  for (let i = 0; i < n; i++) {
    keys.push(istanbulDateKey(new Date(now.getTime() - i * 24 * 60 * 60 * 1000)));
  }
  return keys;
}
