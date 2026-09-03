// Bayi kayıtları — bayiler/<slug>.json ; hızlı giriş için bayiler/_index.json
import { randomBytes, scryptSync, timingSafeEqual } from "crypto";
import { readJson, writeJson, deleteJson, listPaths } from "./store";
import { normalizePricing, type DealerPricing } from "@/data/pricing";

import { SUBSCRIPTION_LABELS, type SubscriptionStatus } from "@/data/subscription";
export { SUBSCRIPTION_LABELS, type SubscriptionStatus };

export interface Subscription {
  status: SubscriptionStatus;
  paidUntil?: string; // YYYY-MM-DD — bu tarihe kadar ödenmiş
  monthlyFee?: number; // TL
  note?: string; // ör. "Son 3 ay alım 62.000 TL — muaf"
}

export interface Dealer {
  slug: string; // bayi kodu — URL ve dosya adı
  username: string;
  passwordHash: string;
  name: string; // firma adı (PDF ve WhatsApp'ta görünür)
  contactName?: string;
  phone: string;
  email?: string;
  address?: string;
  city?: string;
  website?: string;
  active: boolean;
  subscription: Subscription;
  createdAt: string;
  updatedAt: string;
}

export interface DealerIndexRow {
  slug: string;
  username: string;
  name: string;
  active: boolean;
}

const dealerPath = (slug: string) => `bayiler/${slug}.json`;
const pricingPath = (slug: string) => `bayiler/${slug}/fiyat.json`;
const INDEX_PATH = "bayiler/_index.json";

// ---- Şifre ----
export function hashPassword(password: string): string {
  const salt = randomBytes(16).toString("hex");
  const hash = scryptSync(password, salt, 64).toString("hex");
  return `scrypt$${salt}$${hash}`;
}

export function verifyPassword(password: string, stored: string): boolean {
  const [algo, salt, hash] = String(stored || "").split("$");
  if (algo !== "scrypt" || !salt || !hash) return false;
  const test = scryptSync(password, salt, 64);
  const ref = Buffer.from(hash, "hex");
  return test.length === ref.length && timingSafeEqual(test, ref);
}

// ---- Slug / kullanıcı adı normalizasyonu ----
const TR: Record<string, string> = {
  ç: "c", Ç: "c", ğ: "g", Ğ: "g", ı: "i", İ: "i", ö: "o", Ö: "o", ş: "s", Ş: "s", ü: "u", Ü: "u",
};
export function slugify(s: string): string {
  return String(s || "")
    .replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR[c])
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40);
}
export function normalizeUsername(s: string): string {
  return String(s || "")
    .replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR[c])
    .toLowerCase()
    .replace(/\s+/g, "");
}

// ---- Abonelik kuralı ----
/** Bayi panele girebilir mi? (pasif bayi giremez; askıdaki bayi girer ama sipariş açamaz) */
export function dealerCanLogin(d: Dealer): boolean {
  return d.active;
}

/** Bayi yeni sipariş / teklif oluşturabilir mi? */
export function dealerCanOrder(d: Dealer): { ok: boolean; reason?: string } {
  if (!d.active) return { ok: false, reason: "Bayi hesabı pasif." };
  const s = d.subscription;
  if (s.status === "askida") {
    return { ok: false, reason: "Aboneliğiniz askıda. Lütfen Olga Çerçeve ile iletişime geçin." };
  }
  if (s.status === "odeme_bekliyor" && s.paidUntil) {
    // Ödeme gecikmesinde 7 gün tolerans
    const limit = new Date(s.paidUntil);
    limit.setDate(limit.getDate() + 7);
    if (Date.now() > limit.getTime()) {
      return { ok: false, reason: "Abonelik ödemesi gecikti; yeni sipariş oluşturulamıyor." };
    }
  }
  return { ok: true };
}

// ---- CRUD ----
export async function listDealerIndex(): Promise<DealerIndexRow[]> {
  return (await readJson<DealerIndexRow[]>(INDEX_PATH)) || [];
}

async function saveIndex(rows: DealerIndexRow[]): Promise<void> {
  await writeJson(INDEX_PATH, rows);
}

/** İndeksi bayi dosyalarından yeniden kurar (indeks bozulursa). */
export async function rebuildDealerIndex(): Promise<DealerIndexRow[]> {
  const paths = (await listPaths("bayiler/")).filter((p) => /^bayiler\/[a-z0-9-]+\.json$/.test(p));
  const rows: DealerIndexRow[] = [];
  for (const p of paths) {
    const d = await readJson<Dealer>(p);
    if (d) rows.push({ slug: d.slug, username: d.username, name: d.name, active: d.active });
  }
  await saveIndex(rows);
  return rows;
}

export async function getDealer(slug: string): Promise<Dealer | null> {
  if (!/^[a-z0-9-]{2,40}$/.test(slug)) return null;
  return readJson<Dealer>(dealerPath(slug));
}

export async function findDealerByUsername(username: string): Promise<Dealer | null> {
  const q = normalizeUsername(username);
  const rows = await listDealerIndex();
  const hit = rows.find((r) => normalizeUsername(r.username) === q || r.slug === q);
  if (!hit) return null;
  return getDealer(hit.slug);
}

export async function saveDealer(d: Dealer): Promise<boolean> {
  d.updatedAt = new Date().toISOString();
  const ok = await writeJson(dealerPath(d.slug), d);
  if (!ok) return false;
  const rows = await listDealerIndex();
  const row: DealerIndexRow = { slug: d.slug, username: d.username, name: d.name, active: d.active };
  const i = rows.findIndex((r) => r.slug === d.slug);
  if (i >= 0) rows[i] = row;
  else rows.push(row);
  await saveIndex(rows);
  return true;
}

export async function deleteDealer(slug: string): Promise<void> {
  await deleteJson(dealerPath(slug));
  const rows = (await listDealerIndex()).filter((r) => r.slug !== slug);
  await saveIndex(rows);
}

export async function listDealers(): Promise<Dealer[]> {
  const rows = await listDealerIndex();
  const out: Dealer[] = [];
  await Promise.all(
    rows.map(async (r) => {
      const d = await getDealer(r.slug);
      if (d) out.push(d);
    })
  );
  return out.sort((a, b) => a.name.localeCompare(b.name, "tr"));
}

// ---- Fiyatlandırma ----
export async function getDealerPricing(slug: string): Promise<DealerPricing> {
  return normalizePricing(await readJson(pricingPath(slug)));
}

export async function saveDealerPricing(slug: string, raw: unknown): Promise<DealerPricing> {
  const p = normalizePricing(raw);
  p.updatedAt = new Date().toISOString();
  await writeJson(pricingPath(slug), p);
  return p;
}

/** Bayiye ait, istemciye gönderilebilir alanlar (şifre özeti hariç). */
export function publicDealer(d: Dealer) {
  const { passwordHash: _omit, ...rest } = d;
  return rest;
}
export type PublicDealer = ReturnType<typeof publicDealer>;
