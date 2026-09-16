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
  /** İYS filtresi: "0" bilgilendirme, "11" ticari-bireysel, "12" ticari-tacir. */
  iysfilter?: string;
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

export interface SmsHistory {
  records: SmsRecord[]; // en yeni `limit` kayıt, yeniden eskiye
  total: number;        // depodaki toplam gönderim sayısı
}

const SMS_PATH_RE = /^sms\/\d{4}-\d{2}-\d{2}\/[^/]+\.json$/;

/**
 * Son gönderimleri yeniden eskiye getirir.
 *
 * Eski sürüm `list({ prefix: "sms/", limit })` ile depodan "ilk N dosyayı"
 * alıyordu; Blob dosyaları alfabetik (= tarih) sırayla verdiği için bu N,
 * en ESKİ kayıtlar oluyordu ve 200'ü aşınca yeni gönderimler listeye hiç
 * girmiyordu. Şimdi önce tüm kayıt YOLLARI sayfalanarak toplanır (yalnızca
 * meta veri, ucuz), en yeni `limit` tanesi seçilir ve yalnız onlar okunur.
 * Yol = sms/<tarih>/<zamanDamgası-rasgele>.json → alfabetik sıra kronolojiktir.
 */
export async function listSmsHistory(limit = 200): Promise<SmsHistory> {
  if (!blobConfigured()) return { records: [], total: 0 };
  const { list, get } = await import("@vercel/blob");
  const paths: string[] = [];
  try {
    let cursor: string | undefined;
    for (let sayfa = 0; sayfa < 50; sayfa++) {
      const res = await list({ prefix: "sms/", limit: 1000, cursor });
      for (const b of res.blobs) if (SMS_PATH_RE.test(b.pathname)) paths.push(b.pathname);
      if (!res.hasMore || !res.cursor) break;
      cursor = res.cursor;
    }
  } catch {
    /* hiç gönderim yoksa boş liste */
  }
  paths.sort();
  paths.reverse();
  const secilen = paths.slice(0, Math.max(0, limit));
  const out: SmsRecord[] = [];
  await Promise.all(
    secilen.map(async (yol) => {
      try {
        const r = await get(yol, { access: "private", useCache: false });
        if (!r || r.statusCode !== 200 || !r.stream) return;
        out.push(JSON.parse(await new Response(r.stream).text()) as SmsRecord);
      } catch {
        /* tek kayıt okunamazsa listeyi bozma */
      }
    })
  );
  out.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  return { records: out, total: paths.length };
}

/** Geriye dönük uyum: yalnızca kayıt listesi. */
export async function listSmsRecords(limit = 200): Promise<SmsRecord[]> {
  return (await listSmsHistory(limit)).records;
}

/** Çakışmayan, tarihten okunabilir bir kayıt kimliği üretir. */
export function newSmsId(now = new Date()): string {
  const t = now.toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  const r = Math.random().toString(36).slice(2, 7);
  return `${t}-${r}`;
}

export { istanbulDateKey };
