// Personel kartları — finans/personel/<id>.json.
// Avans/maaş ödemeleri ayrı tip DEĞİLDİR: kategori "maaş"/"avans"/"prim" olan
// ve personelId taşıyan Gider kayıtlarıdır. Personel ekranı o giderleri
// kişiye göre süzerek "bu ay çektiği / kalan" görünümünü üretir.
// (Avans-Maaş Excel'inin karşılığı.)

import { blobConfigured } from "./orders";
import type { Branch } from "./customers";

export interface Personel {
  id: string; // P-xxxxx
  name: string;
  branch: Branch;
  startDate?: string; // işe başlama YYYY-MM-DD
  endDate?: string; // işten çıkış
  salary?: number; // aylık maaş (TL)
  note?: string;
  createdAt: string;
  updatedAt: string;
}

const path = (id: string) => `finans/personel/${id}.json`;

export function newPersonelId(): string {
  return "P" + Math.random().toString(36).slice(2, 10).toUpperCase();
}

export async function savePersonel(p: Personel): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(path(p.id), JSON.stringify(p), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function getPersonel(id: string): Promise<Personel | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(path(id), { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as Personel;
  } catch {
    return null;
  }
}

export async function listPersonel(): Promise<Personel[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: Personel[] = [];
  try {
    const { blobs } = await list({ prefix: "finans/personel/", limit: 500 });
    await Promise.all(
      blobs.map(async (b) => {
        try {
          const r = await get(b.pathname, {
            access: "private",
            useCache: false,
          });
          if (!r || r.statusCode !== 200 || !r.stream) return;
          out.push(JSON.parse(await new Response(r.stream).text()) as Personel);
        } catch {
          /* tek kayıt okunamazsa listeyi bozma */
        }
      })
    );
  } catch {
    /* boş */
  }
  return out.sort((a, b) => a.name.localeCompare(b.name, "tr"));
}
