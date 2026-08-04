// Açılış bakiyeleri — finans/acilis/<customerId>.json.
// Excel'den devralınan "geçmişten gelen borç/alacak" tutarı. Cari bakiye:
//   açılış.amount + Σ sipariş(asOf sonrası) − Σ tahsilat(asOf sonrası)
// asOf öncesi hareketler Excel'de kapandığı için tekrar sayılmaz.

import { blobConfigured } from "./orders";
import type { Branch } from "./customers";
import type { FinansKaynak } from "./tahsilat";

export interface AcilisBakiye {
  customerId: string;
  customerName: string; // Excel'deki orijinal ad — mutabakat/denetim için
  branch: Branch;
  amount: number; // TL; pozitif = müşteri borçlu, negatif = alacaklı
  asOf: string; // YYYY-MM-DD — devir tarihi
  note?: string;
  kaynak: FinansKaynak;
  createdAt: string;
  createdBy: string;
}

const path = (customerId: string) => `finans/acilis/${customerId}.json`;

export async function getAcilisBakiye(
  customerId: string
): Promise<AcilisBakiye | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(path(customerId), {
      access: "private",
      useCache: false,
    });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as AcilisBakiye;
  } catch {
    return null;
  }
}

export async function saveAcilisBakiye(a: AcilisBakiye): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(path(a.customerId), JSON.stringify(a), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function listAcilisBakiyeler(): Promise<AcilisBakiye[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: AcilisBakiye[] = [];
  try {
    const { blobs } = await list({ prefix: "finans/acilis/", limit: 1000 });
    await Promise.all(
      blobs.map(async (b) => {
        try {
          const r = await get(b.pathname, {
            access: "private",
            useCache: false,
          });
          if (!r || r.statusCode !== 200 || !r.stream) return;
          out.push(
            JSON.parse(await new Response(r.stream).text()) as AcilisBakiye
          );
        } catch {
          /* tek kayıt okunamazsa listeyi bozma */
        }
      })
    );
  } catch {
    /* boş */
  }
  return out;
}
