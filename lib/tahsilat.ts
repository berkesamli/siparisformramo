// Tahsilat kayıtları — Vercel Blob'da finans/tahsilat/YYYY-MM/<id>.json.
// Sipariş üzerindeki paidAmount artık yalnızca bu modül üzerinden değişir;
// her tahsilat tarih/tutar/yöntem/şube bilgisiyle kalıcı bir harekettir.
//
// Ay önekli yol sayesinde tarih aralığı sorguları yalnızca ilgili ayları
// listeler — tüm geçmişi taramaz (Blob işlem tutumluluğu).

import { blobConfigured, istanbulDateKey } from "./orders";
import type { Branch } from "./customers";

export type TahsilatYontem =
  | "nakit"
  | "havale"
  | "krediKarti"
  | "cek"
  | "senet"
  | "diger";

export const TAHSILAT_YONTEM_LABELS: Record<TahsilatYontem, string> = {
  nakit: "Nakit",
  havale: "Havale / EFT",
  krediKarti: "Kredi Kartı",
  cek: "Çek",
  senet: "Senet",
  diger: "Diğer",
};

// İstanbul kasası YTL dışında USD/EUR nakit de tutuyor (kasa Excel'inden).
export type ParaBirimi = "TL" | "USD" | "EUR";

export type FinansKaynak = "panel" | "migrasyon" | "excel";

export interface Tahsilat {
  id: string; // T-YYYYMMDDHHMMSS-xxxxx
  dateKey: string; // tahsilat tarihi YYYY-MM-DD (İstanbul)
  createdAt: string; // ISO — kayıt anı
  createdBy: string; // kaydeden çalışan
  branch: Branch;
  customerId?: string; // müşteri defteri kaydı (varsa)
  customerName: string; // görünen ad — defter dışı müşteriler için de dolu
  orderId?: string; // belirli siparişe bağlıysa
  orderDateKey?: string; // siparişin blob yolu için
  amount: number; // kuruş yuvarlı
  currency: ParaBirimi; // varsayılan TL
  method: TahsilatYontem;
  /** Parayı fiilen tahsil eden kişi (prim raporu için) — kaydedenden farklı olabilir. */
  tahsilEden?: string;
  cekSenetId?: string; // yöntem çek/senet ise ilgili kayıt (Faz 2)
  note?: string;
  kaynak: FinansKaynak;
}

export const ayKey = (dateKey: string) => dateKey.slice(0, 7); // YYYY-MM

const path = (t: Pick<Tahsilat, "id" | "dateKey">) =>
  `finans/tahsilat/${ayKey(t.dateKey)}/${t.id}.json`;

export function newTahsilatId(now = new Date()): string {
  const t = now.toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  const r = Math.random().toString(36).slice(2, 7);
  return `T-${t}-${r}`;
}

export async function saveTahsilat(t: Tahsilat): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(path(t), JSON.stringify(t), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function getTahsilat(
  ay: string,
  id: string
): Promise<Tahsilat | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(`finans/tahsilat/${ay}/${id}.json`, {
      access: "private",
      useCache: false,
    });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as Tahsilat;
  } catch {
    return null;
  }
}

export async function deleteTahsilat(ay: string, id: string): Promise<void> {
  if (!blobConfigured()) return;
  const { del } = await import("@vercel/blob");
  await del(`finans/tahsilat/${ay}/${id}.json`).catch(() => {});
}

/** Verilen ayların (YYYY-MM) tüm tahsilatları, yeniden eskiye. */
export async function listTahsilatByMonths(
  months: string[]
): Promise<Tahsilat[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: Tahsilat[] = [];
  await Promise.all(
    months.map(async (ay) => {
      try {
        const { blobs } = await list({
          prefix: `finans/tahsilat/${ay}/`,
          limit: 1000,
        });
        await Promise.all(
          blobs.map(async (b) => {
            try {
              const r = await get(b.pathname, {
                access: "private",
                useCache: false,
              });
              if (!r || r.statusCode !== 200 || !r.stream) return;
              out.push(
                JSON.parse(await new Response(r.stream).text()) as Tahsilat
              );
            } catch {
              /* tek kayıt okunamazsa listeyi bozma */
            }
          })
        );
      } catch {
        /* ay klasörü yoksa geç */
      }
    })
  );
  return out.sort((a, b) => b.dateKey.localeCompare(a.dateKey));
}

/**
 * Tüm tahsilatlar — cari sayfası için. Sipariş taramasıyla aynı istekte
 * kullanılır; başka yerde tam tarama yapılmamalıdır.
 */
export async function listAllTahsilat(limit = 3000): Promise<Tahsilat[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: Tahsilat[] = [];
  try {
    const { blobs } = await list({ prefix: "finans/tahsilat/", limit });
    await Promise.all(
      blobs.map(async (b) => {
        try {
          const r = await get(b.pathname, {
            access: "private",
            useCache: false,
          });
          if (!r || r.statusCode !== 200 || !r.stream) return;
          out.push(JSON.parse(await new Response(r.stream).text()) as Tahsilat);
        } catch {
          /* tek kayıt okunamazsa listeyi bozma */
        }
      })
    );
  } catch {
    /* hiç kayıt yoksa boş */
  }
  return out.sort((a, b) => b.dateKey.localeCompare(a.dateKey));
}

export { istanbulDateKey };
