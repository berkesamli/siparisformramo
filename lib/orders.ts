// Sipariş kayıtları — Vercel Blob'da orders/<tarih>/<siparisNo>.json olarak tutulur.
// Yalnızca sunucu tarafında kullanılır.

import type { OrderLine } from "./notify";
import { kurus } from "./num";

export type OrderStatus = "olusturuldu" | "hazirlaniyor" | "tamamlandi" | "iptal";

export const STATUS_LABELS: Record<OrderStatus, string> = {
  olusturuldu: "Oluşturuldu",
  hazirlaniyor: "Hazırlanıyor",
  tamamlandi: "Tamamlandı",
  iptal: "İptal",
};

/**
 * İptal edilen sipariş silinmez (kaydı ve fişi durur) ama ciroya,
 * raporlara, cari bakiyeye ve maliyet analizine dahil edilmez.
 */
export function siparisIptal(o: { status: OrderStatus }): boolean {
  return o.status === "iptal";
}

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
 * Sipariş "tamamlandı" sayılır: durumu tamamlandı + merkez kontrolünden
 * geçmiş. Bu iki şart sağlanınca sipariş aktif listeden çıkar, "Tamamlanan
 * Siparişler" bölümünde arşivlenir. Ödeme durumu ayrı takip edilir (cari
 * hesapta izlenir), arşivlemeyi engellemez.
 */
export function siparisTamamlandi(o: {
  status: OrderStatus;
  kontrol?: { by: string; at: string };
}): boolean {
  return o.status === "tamamlandi" && !!o.kontrol;
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
  await upsertOrderIndex(order);
  return true;
}

const pad3 = (n: number) => String(n).padStart(3, "0");

/**
 * Yeni siparişi ÇAKIŞMASIZ numarayla kaydeder.
 *
 * Eski akış (sayaç oku → +1 → yaz → kaydet) yarışa açıktı: iki şube aynı
 * anda kaydederse aynı OLG numarası üretilir, ikinci kayıt ilkini sessizce
 * ezerdi. Yeni akışta numara önce orders/no/<id>.json rezervasyon dosyasıyla
 * allowOverwrite:false olarak KAPILIR — aynı numarayı ikinci almak isteyen
 * Blob hatası alır ve sıradaki numarayı dener. Sayaç ancak kayıt başarılı
 * olunca güncellenir; bu yüzden kayıt düşerse numara "bildirimde var,
 * panelde yok" hayalet siparişe dönüşmez (çağıran stored=false görür).
 */
export async function createOrder(
  taslak: Omit<SavedOrder, "orderId">
): Promise<{ orderId: string; stored: boolean }> {
  const year = taslak.dateKey.slice(0, 4);

  if (!blobConfigured()) {
    // Blob yok (yerel geliştirme): zaman damgalı benzersiz numara, kayıt yok
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, "0");
    const rand = Math.random().toString(36).slice(2, 5).toUpperCase();
    return {
      orderId: `OLG-${year}-${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}-${rand}`,
      stored: false,
    };
  }

  const { get, put } = await import("@vercel/blob");
  const counterPath = `orders/counter-${year}.json`;
  let seq = 0;
  try {
    const r = await get(counterPath, { access: "private", useCache: false });
    if (r && r.statusCode === 200 && r.stream) {
      seq = Number(JSON.parse(await new Response(r.stream).text()).seq) || 0;
    }
  } catch {
    /* ilk sipariş — sayaç yok */
  }

  for (let deneme = 0; deneme < 6; deneme++) {
    const aday = seq + 1 + deneme;
    const orderId = `OLG-${year}-${pad3(aday)}`;
    try {
      // Numara rezervasyonu — atomik kilit görevi görür
      await put(
        `orders/no/${orderId}.json`,
        JSON.stringify({ orderId, dateKey: taslak.dateKey }),
        { access: "private", contentType: "application/json", addRandomSuffix: false, allowOverwrite: false }
      );
    } catch {
      continue; // numara başka kayıt tarafından kapılmış — sıradakini dene
    }

    const order: SavedOrder = { ...taslak, orderId };
    try {
      await put(orderPath(order.dateKey, orderId), JSON.stringify(order), {
        access: "private", contentType: "application/json", addRandomSuffix: false, allowOverwrite: false,
      });
    } catch {
      // Rezervasyon bizim ama kayıt yazılamadı — bir kez daha dene
      try {
        await put(orderPath(order.dateKey, orderId), JSON.stringify(order), {
          access: "private", contentType: "application/json", addRandomSuffix: false, allowOverwrite: true,
        });
      } catch {
        return { orderId, stored: false };
      }
    }

    try {
      await put(counterPath, JSON.stringify({ seq: aday }), {
        access: "private", contentType: "application/json", addRandomSuffix: false, allowOverwrite: true,
      });
    } catch {
      /* sayaç yazılamazsa sonraki kayıt rezervasyon çakışmasıyla ilerler */
    }
    await upsertOrderIndex(order);
    return { orderId, stored: true };
  }
  return { orderId: "", stored: false };
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

// ---- Aylık sipariş indeksi ----
// Arama (?q=), mükerrer uyarısı (?musteri=) ve cari ekranı her seferinde tüm
// sipariş dosyalarını tek tek okumasın diye ay başına tek indeks dosyası
// tutulur: orders/index/YYYY-MM.json. İndeks bir ÖNBELLEKTİR — gerçek veri
// her zaman orders/<gün>/<no>.json; indeks boşsa/eskiyse tam taramaya düşülür
// ve indeks o taramadan yeniden kurulur.
export interface OrderIndexEntry {
  orderId: string;
  dateKey: string;
  createdAt: string;
  status: OrderStatus;
  employee: string;
  customer: string;
  customerId?: string;
  branch?: "ankara" | "istanbul";
  payment?: PaymentStatus;
  paidAmount?: number;
  net: number;
  note: string;
  // Merkez kontrolü yapıldı mı? Eski indeks kayıtlarında yoktur (undefined):
  // gösterge paneli yalnızca false olanları "kontrol bekliyor" sayar.
  kontrol?: boolean;
}

const indexPath = (ay: string) => `orders/index/${ay}.json`;
const ayOf = (dateKey: string) => dateKey.slice(0, 7);

export function toIndexEntry(o: SavedOrder): OrderIndexEntry {
  return {
    orderId: o.orderId,
    dateKey: o.dateKey,
    createdAt: o.createdAt,
    status: o.status,
    employee: o.employee,
    customer: o.customer,
    customerId: o.customerId || undefined,
    branch: o.branch,
    payment: o.payment,
    paidAmount: o.paidAmount,
    net: o.net,
    note: (o.note || "").slice(0, 200),
    kontrol: Boolean(o.kontrol),
  };
}

/**
 * Siparişi ait olduğu ayın indeksine ekler/günceller. En iyi çaba: iki kayıt
 * tam aynı anda yazarsa biri kaybolabilir — indeks önbellek olduğu için bir
 * sonraki tam tarama (veya yeniden kurma) bunu kendiliğinden düzeltir.
 */
export async function upsertOrderIndex(o: SavedOrder): Promise<void> {
  if (!blobConfigured()) return;
  try {
    const { get, put } = await import("@vercel/blob");
    const yol = indexPath(ayOf(o.dateKey));
    let entries: OrderIndexEntry[] = [];
    try {
      const r = await get(yol, { access: "private", useCache: false });
      if (r && r.statusCode === 200 && r.stream) {
        const parsed = JSON.parse(await new Response(r.stream).text());
        if (Array.isArray(parsed)) entries = parsed;
      }
    } catch {
      /* ilk kayıt — indeks dosyası yok */
    }
    const e = toIndexEntry(o);
    const i = entries.findIndex((x) => x.orderId === o.orderId);
    if (i >= 0) entries[i] = e;
    else entries.push(e);
    await put(yol, JSON.stringify(entries), {
      access: "private",
      contentType: "application/json",
      addRandomSuffix: false,
      allowOverwrite: true,
    });
  } catch {
    /* indeks yazılamazsa arama tam taramaya düşer — sipariş kaydı etkilenmez */
  }
}

/** Son `sonAy` ayın (verilmezse tüm ayların) indeks kayıtlarını getirir. */
export async function readOrderIndex(sonAy?: number): Promise<OrderIndexEntry[]> {
  if (!blobConfigured()) return [];
  try {
    const { list, get } = await import("@vercel/blob");
    const { blobs } = await list({ prefix: "orders/index/", limit: 1000 });
    let files = blobs
      .map((b) => b.pathname)
      .filter((p) => /^orders\/index\/\d{4}-\d{2}\.json$/.test(p))
      .sort()
      .reverse();
    if (sonAy && sonAy > 0) files = files.slice(0, sonAy);
    const out: OrderIndexEntry[] = [];
    await Promise.all(
      files.map(async (p) => {
        try {
          const r = await get(p, { access: "private", useCache: false });
          if (!r || r.statusCode !== 200 || !r.stream) return;
          const arr = JSON.parse(await new Response(r.stream).text());
          if (Array.isArray(arr)) out.push(...arr);
        } catch {
          /* tek ay okunamazsa kalanlarla devam */
        }
      })
    );
    return out.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  } catch {
    return [];
  }
}

/** İndeksi verilen sipariş listesinden baştan kurar; yazılan ay sayısını döner. */
export async function rebuildOrderIndexFrom(orders: SavedOrder[]): Promise<number> {
  if (!blobConfigured() || !orders.length) return 0;
  const { put } = await import("@vercel/blob");
  const aylar = new Map<string, OrderIndexEntry[]>();
  for (const o of orders) {
    const ay = ayOf(o.dateKey);
    if (!aylar.has(ay)) aylar.set(ay, []);
    aylar.get(ay)!.push(toIndexEntry(o));
  }
  let yazilan = 0;
  for (const [ay, entries] of aylar) {
    try {
      await put(indexPath(ay), JSON.stringify(entries), {
        access: "private",
        contentType: "application/json",
        addRandomSuffix: false,
        allowOverwrite: true,
      });
      yazilan++;
    } catch {
      /* tek ay yazılamazsa kalanlarla devam */
    }
  }
  return yazilan;
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
