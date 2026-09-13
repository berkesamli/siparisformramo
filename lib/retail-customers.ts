// Perakende (mağaza) müşteri defteri — Vercel Blob'da
// retail-customers/<id>.json olarak tutulur. Etiket/toptan defterinden
// (customers/) BİLEREK ayrıdır: mağaza müşterisi son tüketicidir, firma,
// şube, kargo etiketi gibi alanları yoktur; sipariş kaydında telefon
// üzerinden kendiliğinden oluşur/güncellenir.

import { blobConfigured } from "./orders";

export interface RetailCustomer {
  id: string; // P ile başlar (perakende)
  name: string;
  phone: string;
  email: string;
  address: string;
  note: string;
  createdAt: string;
  updatedAt: string;
}

const path = (id: string) => `retail-customers/${id}.json`;

const newId = () => "P" + Math.random().toString(36).slice(2, 10).toUpperCase();

const s = (v: unknown, max = 160) => String(v ?? "").trim().slice(0, max);

/** Telefonları karşılaştırma anahtarı: yalnız rakamlar, son 10 hane. */
export function phoneKey(phone: string): string {
  const d = String(phone || "").replace(/\D/g, "");
  return d.length > 10 ? d.slice(-10) : d;
}

export function sanitizeRetailCustomer(
  raw: any,
  existing?: RetailCustomer
): RetailCustomer {
  const now = new Date().toISOString();
  return {
    id: existing?.id || s(raw?.id, 40) || newId(),
    name: s(raw?.name, 120),
    phone: s(raw?.phone, 40),
    email: s(raw?.email, 120),
    address: s(raw?.address, 300),
    note: s(raw?.note, 300),
    createdAt: existing?.createdAt || now,
    updatedAt: now,
  };
}

export async function saveRetailCustomer(c: RetailCustomer): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(path(c.id), JSON.stringify(c), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function getRetailCustomer(id: string): Promise<RetailCustomer | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(path(id), { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as RetailCustomer;
  } catch {
    return null;
  }
}

export async function listRetailCustomers(): Promise<RetailCustomer[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: RetailCustomer[] = [];
  try {
    const { blobs } = await list({ prefix: "retail-customers/", limit: 1000 });
    await Promise.all(
      blobs.map(async (b) => {
        try {
          const r = await get(b.pathname, { access: "private", useCache: false });
          if (!r || r.statusCode !== 200 || !r.stream) return;
          out.push(JSON.parse(await new Response(r.stream).text()) as RetailCustomer);
        } catch {
          /* tek kayıt okunamazsa listeyi bozma */
        }
      })
    );
  } catch {
    return [];
  }
  return out.sort((a, b) =>
    a.name.localeCompare(b.name, "tr", { sensitivity: "base" })
  );
}

export async function deleteRetailCustomer(id: string): Promise<boolean> {
  if (!blobConfigured()) return false;
  try {
    const { del } = await import("@vercel/blob");
    await del(path(id));
    return true;
  } catch {
    return false;
  }
}

/**
 * Sipariş kaydında müşteriyi deftere kendiliğinden işler: telefon eşleşen
 * kayıt varsa boş alanları tamamlar, yoksa yeni kayıt açar. Dönen id
 * siparişin customerId alanına yazılır. En iyi çaba — defter yazılamazsa
 * sipariş etkilenmez.
 */
export async function upsertRetailCustomerFromOrder(bilgi: {
  name: string;
  phone: string;
  email?: string;
  address?: string;
}): Promise<string | null> {
  if (!blobConfigured()) return null;
  const key = phoneKey(bilgi.phone);
  if (!bilgi.name || key.length < 7) return null;
  try {
    const hepsi = await listRetailCustomers();
    const mevcut = hepsi.find((c) => phoneKey(c.phone) === key);
    if (mevcut) {
      const guncel: RetailCustomer = {
        ...mevcut,
        name: bilgi.name || mevcut.name,
        email: mevcut.email || s(bilgi.email, 120),
        address: mevcut.address || s(bilgi.address, 300),
        updatedAt: new Date().toISOString(),
      };
      await saveRetailCustomer(guncel);
      return guncel.id;
    }
    const yeni = sanitizeRetailCustomer({
      name: bilgi.name,
      phone: bilgi.phone,
      email: bilgi.email,
      address: bilgi.address,
    });
    await saveRetailCustomer(yeni);
    return yeni.id;
  } catch {
    return null;
  }
}
