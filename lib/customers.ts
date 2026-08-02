// Müşteri defteri — Vercel Blob'da customers/<id>.json olarak tutulur.
// Hem kargo etiketi hem de toptan sipariş formundaki müşteri seçici kullanır.
// (Eski "etiket" uygulamasındaki Google Sheets tablosunun karşılığı.)

import { blobConfigured } from "./orders";

export interface Customer {
  id: string;
  firstName: string; // kişi adı veya firma adı
  lastName: string;
  company: string; // firma ünvanı (opsiyonel)
  email: string;
  phone: string;
  addr1: string;
  addr2: string;
  city: string; // İl
  district: string; // İlçe
  postalCode: string;
  country: string;
  branch: Branch; // hangi şubeden gönderiliyor
  note: string;
  createdAt: string;
  updatedAt: string;
}

export type Branch = "ankara" | "istanbul";

export interface BranchInfo {
  id: Branch;
  label: string;
  name: string;
  cityTel: string;
  addr1: string;
  addr2: string;
  website: string;
}

// Gönderici şubeler — etikette "Gönderici" alanı buna göre değişir.
export const BRANCHES: Record<Branch, BranchInfo> = {
  ankara: {
    id: "ankara",
    label: "Ankara (Merkez)",
    name: "Olga Çerçeve",
    cityTel: "Ankara • Tel: 0312 495 75 45",
    addr1: "Birlik Mahallesi 448. Cadde No:56",
    addr2: "Çankaya / Ankara",
    website: "olgacerceve.com",
  },
  istanbul: {
    id: "istanbul",
    label: "İstanbul",
    name: "Olga Çerçeve",
    cityTel: "İstanbul • Tel: 0212 675 27 50",
    addr1: "Masko Mobilya Sanayi Sitesi, 3-B1 No:4",
    addr2: "İkitelli / Başakşehir, 34490 İstanbul",
    website: "olgacerceve.com",
  },
};

export function branchInfo(b: unknown): BranchInfo {
  return b === "istanbul" ? BRANCHES.istanbul : BRANCHES.ankara;
}

/** Müşterinin listelerde/etikette görünen tam adı (firma öncelikli). */
export function customerTitle(c: Pick<Customer, "company" | "firstName" | "lastName">): string {
  const person = `${c.firstName || ""} ${c.lastName || ""}`.trim();
  if (c.company && person) return `${c.company} — ${person}`;
  return c.company || person || "-";
}

const path = (id: string) => `customers/${id}.json`;

const newId = () => "C" + Math.random().toString(36).slice(2, 10).toUpperCase();

const s = (v: unknown, max = 160) => String(v ?? "").trim().slice(0, max);

export function sanitizeCustomer(raw: any, existing?: Customer): Customer {
  const now = new Date().toISOString();
  return {
    id: existing?.id || s(raw?.id, 40) || newId(),
    firstName: s(raw?.firstName, 80),
    lastName: s(raw?.lastName, 80),
    company: s(raw?.company, 120),
    email: s(raw?.email, 120),
    phone: s(raw?.phone, 40),
    addr1: s(raw?.addr1, 200),
    addr2: s(raw?.addr2, 200),
    city: s(raw?.city, 60),
    district: s(raw?.district, 60),
    postalCode: s(raw?.postalCode, 20),
    country: s(raw?.country, 60) || "Türkiye",
    branch: raw?.branch === "istanbul" ? "istanbul" : "ankara",
    note: s(raw?.note, 300),
    createdAt: existing?.createdAt || now,
    updatedAt: now,
  };
}

export async function saveCustomer(c: Customer): Promise<boolean> {
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

export async function getCustomer(id: string): Promise<Customer | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(path(id), { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as Customer;
  } catch {
    return null;
  }
}

export async function listCustomers(): Promise<Customer[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: Customer[] = [];
  try {
    const { blobs } = await list({ prefix: "customers/", limit: 1000 });
    await Promise.all(
      blobs.map(async (b) => {
        try {
          const r = await get(b.pathname, { access: "private", useCache: false });
          if (!r || r.statusCode !== 200 || !r.stream) return;
          out.push(JSON.parse(await new Response(r.stream).text()) as Customer);
        } catch {
          /* tek kayıt okunamazsa listeyi bozma */
        }
      })
    );
  } catch {
    return [];
  }
  return out.sort((a, b) =>
    customerTitle(a).localeCompare(customerTitle(b), "tr", { sensitivity: "base" })
  );
}

export async function deleteCustomer(id: string): Promise<boolean> {
  if (!blobConfigured()) return false;
  try {
    const { del } = await import("@vercel/blob");
    await del(path(id));
    return true;
  } catch {
    return false;
  }
}

/** Şehir adını karşılaştırma için sadeleştirir (İstanbul = istanbul = ISTANBUL). */
export function normalizeCity(s: string): string {
  const TR: Record<string, string> = {
    ç: "c", Ç: "c", ğ: "g", Ğ: "g", ı: "i", İ: "i",
    ö: "o", Ö: "o", ş: "s", Ş: "s", ü: "u", Ü: "u",
  };
  return String(s || "")
    .replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR[c])
    .toLowerCase()
    .trim();
}
