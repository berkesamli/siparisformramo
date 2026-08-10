// Sipariş kayıtları — Vercel Blob'da orders/<tarih>/<siparisNo>.json olarak tutulur.
// Yalnızca sunucu tarafında kullanılır.

import type { OrderLine } from "./notify";
import { kurus } from "./num";

export type OrderStatus = "olusturuldu" | "hazirlaniyor" | "tamamlandi";

export const STATUS_LABELS: Record<OrderStatus, string> = {
  olusturuldu: "Oluşturuldu",
  hazirlaniyor: "Hazırlanıyor",
  tamamlandi: "Tamamlandı",
};

// ---- Ödeme (cari) takibi ----
export type PaymentStatus = "bekliyor" | "kismi" | "odendi";

export const PAYMENT_LABELS: Record<PaymentStatus, string> = {
  bekliyor: "Ödeme Bekliyor",
  kismi: "Kısmi Ödendi",
  odendi: "Ödendi",
};

/** Siparişin kalan bakiyesi (net − tahsil edilen). */
export function orderBalance(o: {
  net: number;
  payment?: PaymentStatus;
  paidAmount?: number;
}): number {
  if (o.payment === "odendi") return 0;
  const paid = Number(o.paidAmount) || 0;
  return Math.max(0, Math.round((o.net - paid) * 100) / 100);
}

/**
 * Sipariş "tamamlandı" sayılır: üretimi bitmiş + ödemesi alınmış + merkez
 * kontrolünden geçmiş. Bu üçü tamamlanınca sipariş aktif listeden çıkar,
 * "Tamamlanan Siparişler" bölümünde arşivlenir.
 */
export function siparisTamamlandi(o: {
  status: OrderStatus;
  payment?: PaymentStatus;
  kontrol?: { by: string; at: string };
}): boolean {
  return o.status === "tamamlandi" && o.payment === "odendi" && !!o.kontrol;
}

export interface SavedOrder {
  orderId: string;
  dateKey: string; // YYYY-MM-DD (İstanbul)
  createdAt: string; // ISO
  updatedAt: string; // ISO
  status: OrderStatus;
  employee: string;
  customer: string;
  customerId?: string; // müşteri defterindeki kayıt (varsa)
  // Siparişin şubesi — formda seçilir, müşterinin şubesi önerilir. Eski
  // kayıtlarda yoktur; raporlar müşteri eşleşmesinden türetir, yoksa "belirsiz".
  branch?: "ankara" | "istanbul";
  payment?: PaymentStatus;
  paidAmount?: number; // tahsil edilen tutar (kısmi ödemede)
  // Merkez kontrolü: mesai sonrası (19:00+) girilen siparişler ertesi gün
  // gözden kaçmasın diye her sipariş tek tek "kontrol edildi" işaretlenir.
  kontrol?: { by: string; at: string };
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
  // Satır tutarları zaten kuruşa yuvarlanmış gelir; burada yalnızca
  // toplama/iskonto/KDV adımları kuruşa sabitlenir. Bkz. lib/num.ts
  const r2 = kurus;
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

/** Tüm toptan siparişleri getirir (müşteri geçmişi ve raporlar için). */
export async function listAllOrders(limit = 2000): Promise<SavedOrder[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const orders: SavedOrder[] = [];
  try {
    const { blobs } = await list({ prefix: "orders/", limit });
    await Promise.all(
      blobs
        // sayaç dosyalarını atla: orders/counter-2026.json
        .filter((b) => /^orders\/\d{4}-\d{2}-\d{2}\//.test(b.pathname))
        .map(async (b) => {
          try {
            const r = await get(b.pathname, { access: "private", useCache: false });
            if (!r || r.statusCode !== 200 || !r.stream) return;
            orders.push(JSON.parse(await new Response(r.stream).text()) as SavedOrder);
          } catch {
            /* tek kayıt okunamazsa listeyi bozma */
          }
        })
    );
  } catch {
    return [];
  }
  return orders.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

/** Sıralı sipariş numarası: OLG-2026-001. Sayaç Blob'da yıl bazlı tutulur. */
export async function nextOrderId(): Promise<string> {
  const now = new Date();
  const year = istanbulDateKey(now).slice(0, 4);

  if (!blobConfigured()) {
    // Blob yoksa zaman damgalı benzersiz numara
    const pad = (n: number) => String(n).padStart(2, "0");
    const rand = Math.random().toString(36).slice(2, 5).toUpperCase();
    return `OLG-${year}-${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}-${rand}`;
  }

  const { get, put } = await import("@vercel/blob");
  const counterPath = `orders/counter-${year}.json`;
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
  return `OLG-${year}-${String(seq).padStart(3, "0")}`;
}

// ---- Günlük kur (ilk giren yazar, gün boyu formlara otomatik gelir) ----

export interface DailyRates {
  rate: number; // TL/USD
  euroRate: number; // TL/EUR
  updatedAt: string;
  by: string;
  // Yetkili (firma sahibi) tarafından belirlendiyse true — bu durumda diğer
  // çalışanların sipariş formunda kur alanı kilitlenir. Kur girilmeden
  // sipariş alınırsa eski davranış sürer (ilk sipariş günün kurunu yazar,
  // ama kimseyi kilitlemez).
  sabit?: boolean;
}

const ratesPath = (dateKey: string) => `rates/${dateKey}.json`;

export async function getDailyRates(dateKey: string): Promise<DailyRates | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(ratesPath(dateKey), { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as DailyRates;
  } catch {
    return null;
  }
}

export async function saveDailyRates(dateKey: string, rates: DailyRates): Promise<void> {
  if (!blobConfigured()) return;
  const { put } = await import("@vercel/blob");
  await put(ratesPath(dateKey), JSON.stringify(rates), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
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
