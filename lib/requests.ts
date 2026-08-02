// Bayi sipariş talepleri — müşteri portalından gelen, çalışan onayı bekleyen
// istekler. Blob'da requests/<YYYY-MM-DD>/<id>.json olarak tutulur.
// Fiyatlandırma çalışan tarafında yapılır; bayi yalnızca ürün ve miktar bildirir.

import { blobConfigured, istanbulDateKey } from "./orders";

export type RequestStatus = "bekliyor" | "onaylandi" | "reddedildi";

export const REQUEST_LABELS: Record<RequestStatus, string> = {
  bekliyor: "Onay Bekliyor",
  onaylandi: "Onaylandı",
  reddedildi: "Reddedildi",
};

export interface RequestLine {
  code: string; // profil / ürün kodu
  unit: string; // Metre | Koli | Adet
  qty: number;
  note: string;
}

export interface SavedRequest {
  id: string;
  dateKey: string;
  createdAt: string;
  updatedAt: string;
  status: RequestStatus;
  username: string; // talebi gönderen portal kullanıcısı
  customer: string; // bayi / firma adı
  phone: string;
  note: string;
  lines: RequestLine[];
  handledBy?: string; // işleme alan çalışan
  orderId?: string; // onaylanıp siparişe dönüştüyse
}

const path = (dateKey: string, id: string) => `requests/${dateKey}/${id}.json`;

const newId = () => "T" + Date.now().toString(36).toUpperCase();

const s = (v: unknown, max = 120) => String(v ?? "").trim().slice(0, max);

export function sanitizeRequestLines(raw: unknown): RequestLine[] {
  if (!Array.isArray(raw)) return [];
  return raw
    .map((l: any) => ({
      code: s(l?.code, 60),
      unit: s(l?.unit, 20) || "Metre",
      qty: Math.max(0, Number(l?.qty) || 0),
      note: s(l?.note, 120),
    }))
    .filter((l) => l.code && l.qty > 0)
    .slice(0, 60);
}

export async function saveRequest(r: SavedRequest): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(path(r.dateKey, r.id), JSON.stringify(r), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function getRequest(dateKey: string, id: string): Promise<SavedRequest | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(path(dateKey, id), { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as SavedRequest;
  } catch {
    return null;
  }
}

/** Tüm talepler (çalışan listesi ve bayinin kendi geçmişi için). */
export async function listRequests(limit = 1000): Promise<SavedRequest[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: SavedRequest[] = [];
  try {
    const { blobs } = await list({ prefix: "requests/", limit });
    await Promise.all(
      blobs.map(async (b) => {
        try {
          const r = await get(b.pathname, { access: "private", useCache: false });
          if (!r || r.statusCode !== 200 || !r.stream) return;
          out.push(JSON.parse(await new Response(r.stream).text()) as SavedRequest);
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

export function newRequest(
  data: Omit<SavedRequest, "id" | "dateKey" | "createdAt" | "updatedAt" | "status">
): SavedRequest {
  const now = new Date();
  return {
    ...data,
    id: newId(),
    dateKey: istanbulDateKey(now),
    createdAt: now.toISOString(),
    updatedAt: now.toISOString(),
    status: "bekliyor",
  };
}
