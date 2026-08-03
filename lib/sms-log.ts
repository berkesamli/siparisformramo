// Gönderilen SMS kayıtları — Vercel Blob'da sms/<tarih>/<id>.json olarak tutulur.
// Kimin ne zaman kime ne gönderdiği ve kaç kredi harcandığı buradan izlenir.

import { blobConfigured, istanbulDateKey } from "./orders";

export interface SmsRecord {
  id: string;
  dateKey: string; // YYYY-MM-DD (İstanbul)
  createdAt: string; // ISO
  sender: string; // gönderen çalışanın adı
  message: string;
  recipients: string[]; // normalize edilmiş numaralar
  segments: number; // mesaj başına SMS parçası
  credits: number; // segments × alıcı sayısı
  ok: boolean;
  jobId?: string;
  error?: string;
}

const path = (dateKey: string, id: string) => `sms/${dateKey}/${id}.json`;

export async function saveSmsRecord(rec: SmsRecord): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(path(rec.dateKey, rec.id), JSON.stringify(rec), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

/** Son gönderimleri yeniden eskiye getirir. */
export async function listSmsRecords(limit = 200): Promise<SmsRecord[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: SmsRecord[] = [];
  try {
    const { blobs } = await list({ prefix: "sms/", limit });
    await Promise.all(
      blobs.map(async (b) => {
        try {
          const r = await get(b.pathname, { access: "private", useCache: false });
          if (!r || r.statusCode !== 200 || !r.stream) return;
          out.push(JSON.parse(await new Response(r.stream).text()) as SmsRecord);
        } catch {
          /* tek kayıt okunamazsa listeyi bozma */
        }
      })
    );
  } catch {
    /* hiç gönderim yoksa boş liste */
  }
  return out.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

/** Çakışmayan, tarihten okunabilir bir kayıt kimliği üretir. */
export function newSmsId(now = new Date()): string {
  const t = now.toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  const r = Math.random().toString(36).slice(2, 7);
  return `${t}-${r}`;
}

export { istanbulDateKey };
