// Gider kayıtları — finans/gider/YYYY-MM/<id>.json (ay önekli, tahsilat gibi).
// Kategoriler İstanbul kasa Excel'indeki gerçek listeden geliyor; serbest
// metin de girilebilir.

import { blobConfigured, istanbulDateKey } from "./orders";
import type { Branch } from "./customers";
import type { FinansKaynak, ParaBirimi } from "./tahsilat";

export type GiderYontem = "nakit" | "havale" | "krediKarti" | "cek" | "diger";

export const GIDER_YONTEM_LABELS: Record<GiderYontem, string> = {
  nakit: "Nakit",
  havale: "Havale / EFT",
  krediKarti: "Kredi Kartı",
  cek: "Çek (verilen)",
  diger: "Diğer",
};

// İstanbul kasa analizindeki çıkış kalemleri + genel ihtiyaçlar.
export const GIDER_KATEGORILERI = [
  "satıcılar",
  "kira",
  "aidat",
  "elektrik",
  "su",
  "doğalgaz",
  "internet-telefon",
  "maaş",
  "avans",
  "prim",
  "yemek",
  "nakliye",
  "malzeme",
  "vergi-sgk",
  "banka masrafı",
  "muhtelif",
] as const;

export interface Gider {
  id: string; // G-YYYYMMDDHHMMSS-xxxxx
  dateKey: string; // YYYY-MM-DD
  createdAt: string;
  createdBy: string;
  branch: Branch;
  category: string;
  description: string;
  amount: number;
  currency: ParaBirimi;
  method: GiderYontem;
  supplier?: string; // ödenen taraf / tedarikçi
  personelId?: string; // maaş/avans ödemesiyse ilgili personel kartı
  cekSenetId?: string; // çekle ödendiyse verilen çek kaydı
  note?: string;
  kaynak: FinansKaynak;
}

const path = (g: Pick<Gider, "id" | "dateKey">) =>
  `finans/gider/${g.dateKey.slice(0, 7)}/${g.id}.json`;

export function newGiderId(now = new Date()): string {
  const t = now.toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  return `G-${t}-${Math.random().toString(36).slice(2, 7)}`;
}

export async function saveGider(g: Gider): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(path(g), JSON.stringify(g), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function getGider(ay: string, id: string): Promise<Gider | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(`finans/gider/${ay}/${id}.json`, {
      access: "private",
      useCache: false,
    });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as Gider;
  } catch {
    return null;
  }
}

export async function deleteGider(ay: string, id: string): Promise<void> {
  if (!blobConfigured()) return;
  const { del } = await import("@vercel/blob");
  await del(`finans/gider/${ay}/${id}.json`).catch(() => {});
}

export async function listGiderByMonths(months: string[]): Promise<Gider[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: Gider[] = [];
  await Promise.all(
    months.map(async (ay) => {
      try {
        const { blobs } = await list({
          prefix: `finans/gider/${ay}/`,
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
              out.push(JSON.parse(await new Response(r.stream).text()) as Gider);
            } catch {
              /* tek kayıt okunamazsa listeyi bozma */
            }
          })
        );
      } catch {
        /* ay yoksa geç */
      }
    })
  );
  return out.sort((a, b) => b.dateKey.localeCompare(a.dateKey));
}

export { istanbulDateKey };
