// Perakende (online çerçeve) siparişleri — Vercel Blob'da
// retail/<tarih>/<siparisNo>.json olarak tutulur. Yalnızca sunucu tarafında kullanılır.

import { findProfile } from "@/data/catalog";
import type { RetailStatus } from "@/data/perakende";
import { blobConfigured, istanbulDateKey, type PaymentStatus } from "./orders";

// Perakende metre fiyatı = katalog liste fiyatı (USD/mt) × PERAKENDE_KATSAYI × kur.
// Katsayı gerekirse RETAIL_FACTOR env değişkeniyle değiştirilebilir.
const RETAIL_FACTOR = Number(process.env.RETAIL_FACTOR) || 5;

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
  // Pencereli (çoklu açıklık) paspartu — websitedeki hesaplayıcıyla aynı:
  // ölçü tek fotoğrafın ölçüsüdür, alan pencere düzeninden türetilir.
  pencereSayisi?: number; // 1..9 (yoksa 1)
  pencereDuzen?: string; // "2x3" (satır x sütun)
  pencereAralik?: number; // pencereler arası görünen köprü (mm)
  // Kasa (kanvas) çerçeve: cam ve paspartu uygulanmaz, tuval kasaya gerilir
  kasa?: boolean;
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
  customerAddress?: string;
  customerId?: string; // müşteri defterindeki kayıt (varsa)
  branch?: "ankara" | "istanbul";
  // Perakendede KDV kavramı olmadığı için faturalı işareti ayrı tutulur
  // (toptanda faturalı = vatApplied'dan türetilir).
  faturali?: boolean;
  payment?: PaymentStatus;
  paidAmount?: number;
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
  await upsertRetailIndex(order);
  return true;
}

/**
 * Yeni perakende siparişi ÇAKIŞMASIZ numarayla kaydeder — toptandaki
 * createOrder ile aynı desen: numara retail/no/<id>.json rezervasyonuyla
 * (allowOverwrite:false) kapılır, sayaç ancak kayıt başarılı olunca
 * güncellenir. İki personel aynı anda kaydederse ikisi de kendi PRK
 * numarasını alır; hiçbir kayıt diğerini ezmez.
 */
export async function createRetailOrder(
  taslak: Omit<SavedRetailOrder, "orderId">
): Promise<{ orderId: string; stored: boolean }> {
  const year = taslak.dateKey.slice(0, 4);

  if (!blobConfigured()) {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, "0");
    const rand = Math.random().toString(36).slice(2, 5).toUpperCase();
    return {
      orderId: `PRK-${year}-${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}-${rand}`,
      stored: false,
    };
  }

  const { get, put } = await import("@vercel/blob");
  const counterPath = `retail/counter-${year}.json`;
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
    const orderId = `PRK-${year}-${String(aday).padStart(3, "0")}`;
    try {
      await put(
        `retail/no/${orderId}.json`,
        JSON.stringify({ orderId, dateKey: taslak.dateKey }),
        { access: "private", contentType: "application/json", addRandomSuffix: false, allowOverwrite: false }
      );
    } catch {
      continue; // numara kapılmış — sıradakini dene
    }
    const order = { ...taslak, orderId } as SavedRetailOrder;
    try {
      await put(orderPath(order.dateKey, orderId), JSON.stringify(order), {
        access: "private", contentType: "application/json", addRandomSuffix: false, allowOverwrite: false,
      });
    } catch {
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
      /* sayaç yazılamazsa sonraki kayıt rezervasyonla ilerler */
    }
    await upsertRetailIndex(order);
    return { orderId, stored: true };
  }
  return { orderId: "", stored: false };
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

/** Tüm perakende siparişleri (müşteri geçmişi ve raporlar için). */
export async function listAllRetailOrders(limit = 2000): Promise<SavedRetailOrder[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: SavedRetailOrder[] = [];
  try {
    const { blobs } = await list({ prefix: "retail/", limit });
    await Promise.all(
      blobs
        .filter((b) => /^retail\/\d{4}-\d{2}-\d{2}\//.test(b.pathname))
        .map(async (b) => {
          try {
            const r = await get(b.pathname, { access: "private", useCache: false });
            if (!r || r.statusCode !== 200 || !r.stream) return;
            out.push(JSON.parse(await new Response(r.stream).text()) as SavedRetailOrder);
          } catch {
            /* tek kayıt okunamazsa listeyi bozma */
          }
        })
    );
  } catch {
    return [];
  }
  return out.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
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


// ---- Aylık perakende indeksi ----
// Gösterge paneli ve genel arama her siparişi tek tek okumasın diye ay başına
// tek indeks dosyası: retail/index/YYYY-MM.json. İndeks bir ÖNBELLEKTİR —
// gerçek veri her zaman retail/<gün>/<no>.json; indeks boşsa tam taramaya
// düşülür ve o taramadan yeniden kurulur (toptan indeksiyle aynı desen).
export interface RetailIndexEntry {
  orderId: string;
  dateKey: string;
  createdAt: string;
  status: RetailStatus;
  employee: string;
  customerName: string;
  customerPhone: string;
  customerId?: string;
  branch?: "ankara" | "istanbul";
  payment?: PaymentStatus;
  paidAmount?: number;
  total: number;
  deliveryDate: string;
  adet: number; // kalem sayısı
}

const retailIndexPath = (ay: string) => `retail/index/${ay}.json`;
const ayOfKey = (dateKey: string) => dateKey.slice(0, 7);

export function toRetailIndexEntry(o: SavedRetailOrder): RetailIndexEntry {
  return {
    orderId: o.orderId,
    dateKey: o.dateKey,
    createdAt: o.createdAt,
    status: o.status,
    employee: o.employee,
    customerName: o.customerName,
    customerPhone: o.customerPhone,
    customerId: o.customerId || undefined,
    branch: o.branch,
    payment: o.payment,
    paidAmount: o.paidAmount,
    total: o.total,
    deliveryDate: o.deliveryDate,
    adet: Array.isArray(o.items) ? o.items.length : 0,
  };
}

export async function upsertRetailIndex(o: SavedRetailOrder): Promise<void> {
  if (!blobConfigured()) return;
  try {
    const { get, put } = await import("@vercel/blob");
    const yol = retailIndexPath(ayOfKey(o.dateKey));
    let entries: RetailIndexEntry[] = [];
    try {
      const r = await get(yol, { access: "private", useCache: false });
      if (r && r.statusCode === 200 && r.stream) {
        const parsed = JSON.parse(await new Response(r.stream).text());
        if (Array.isArray(parsed)) entries = parsed;
      }
    } catch {
      /* ilk kayıt — indeks dosyası yok */
    }
    const e = toRetailIndexEntry(o);
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
    /* indeks yazılamazsa okuyanlar tam taramaya düşer — sipariş kaydı etkilenmez */
  }
}

/** Son `sonAy` ayın (verilmezse tüm ayların) perakende indeks kayıtları. */
export async function readRetailIndex(sonAy?: number): Promise<RetailIndexEntry[]> {
  if (!blobConfigured()) return [];
  try {
    const { list, get } = await import("@vercel/blob");
    const { blobs } = await list({ prefix: "retail/index/", limit: 1000 });
    let files = blobs
      .map((b) => b.pathname)
      .filter((p) => /^retail\/index\/\d{4}-\d{2}\.json$/.test(p))
      .sort()
      .reverse();
    if (sonAy && sonAy > 0) files = files.slice(0, sonAy);
    const out: RetailIndexEntry[] = [];
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
export async function rebuildRetailIndexFrom(orders: SavedRetailOrder[]): Promise<number> {
  if (!blobConfigured() || !orders.length) return 0;
  const { put } = await import("@vercel/blob");
  const aylar = new Map<string, RetailIndexEntry[]>();
  for (const o of orders) {
    const ay = ayOfKey(o.dateKey);
    if (!aylar.has(ay)) aylar.set(ay, []);
    aylar.get(ay)!.push(toRetailIndexEntry(o));
  }
  let yazilan = 0;
  for (const [ay, entries] of aylar) {
    try {
      await put(retailIndexPath(ay), JSON.stringify(entries), {
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

/**
 * Perakende indeksini okur; hiç yoksa (ilk kurulum) tüm siparişleri tarayıp
 * indeksi kurar ve o taramadan döner. Kendi kendini onarır.
 */
export async function readRetailIndexOrRebuild(sonAy?: number): Promise<RetailIndexEntry[]> {
  const idx = await readRetailIndex(sonAy);
  if (idx.length) return idx;
  const hepsi = await listAllRetailOrders();
  if (!hepsi.length) return [];
  await rebuildRetailIndexFrom(hepsi);
  const entries = hepsi.map(toRetailIndexEntry);
  if (sonAy && sonAy > 0) {
    const aylar = [...new Set(entries.map((e) => ayOfKey(e.dateKey)))].sort().reverse().slice(0, sonAy);
    const izin = new Set(aylar);
    return entries.filter((e) => izin.has(ayOfKey(e.dateKey)));
  }
  return entries;
}
