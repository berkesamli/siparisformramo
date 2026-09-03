// Bayi siparişleri — bayiler/<slug>/siparisler/<tarih>/<siparisNo>.json
import { createHmac } from "crypto";
import { readJson, writeJson, listPaths, readMany, istanbulDateKey } from "./store";
import { authSecret } from "./jwt";
import type { OrderStatus, PaymentStatus } from "@/data/pricing";

export interface OrderItem {
  artWidth: number;
  artWidthUnit: "cm" | "mm";
  artHeight: number;
  artHeightUnit: "cm" | "mm";
  frameCode: string;
  framePriceTL: number; // TL/m (bayinin satış fiyatı)
  manualPrice: boolean;
  matType: string;
  matCode: string;
  matColor: string;
  matColorHex: string;
  doubleMat: boolean;
  innerMatType: string;
  innerMatColor: string;
  innerMatColorHex: string;
  altMontaj: string;
  zeminEnabled: boolean;
  zeminType: string;
  zeminColor: string;
  zeminColorHex: string;
  matTop: number;
  matRight: number;
  matBottom: number;
  matLeft: number;
  glassType: string;
  printType: string;
  frameCost: number;
  matCost: number;
  glassCost: number;
  printCost: number;
  laborCost: number;
  itemTotal: number;
}

export interface SavedOrder {
  orderId: string;
  dealerSlug: string;
  dateKey: string; // YYYY-MM-DD (İstanbul)
  createdAt: string;
  updatedAt: string;
  status: OrderStatus;
  createdBy: string;
  customerName: string;
  customerPhone: string;
  customerEmail: string;
  customerAddress?: string;
  payment: PaymentStatus;
  paidAmount: number;
  usdRate: number;
  deliveryDate: string;
  notes: string;
  items: OrderItem[];
  gross: number;
  discount: number;
  total: number;
}

const orderPath = (slug: string, dateKey: string, orderId: string) =>
  `bayiler/${slug}/siparisler/${dateKey}/${orderId}.json`;

export async function saveOrder(o: SavedOrder): Promise<boolean> {
  return writeJson(orderPath(o.dealerSlug, o.dateKey, o.orderId), o);
}

export async function getOrder(
  slug: string,
  dateKey: string,
  orderId: string
): Promise<SavedOrder | null> {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey) || !/^[A-Z0-9-]+$/i.test(orderId)) return null;
  const o = await readJson<SavedOrder>(orderPath(slug, dateKey, orderId));
  return o && o.dealerSlug === slug ? o : null;
}

export async function listOrders(slug: string, dateKeys: string[]): Promise<SavedOrder[]> {
  const paths: string[] = [];
  await Promise.all(
    dateKeys.map(async (k) => {
      paths.push(...(await listPaths(`bayiler/${slug}/siparisler/${k}/`, 500)));
    })
  );
  const orders = await readMany<SavedOrder>(paths);
  return orders.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function listAllOrders(slug: string, limit = 1000): Promise<SavedOrder[]> {
  const paths = (await listPaths(`bayiler/${slug}/siparisler/`, limit)).filter((p) =>
    /\/siparisler\/\d{4}-\d{2}-\d{2}\//.test(p)
  );
  const orders = await readMany<SavedOrder>(paths);
  return orders.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

/** Bayi başına sıralı sipariş numarası: SIP-2026-001 */
export async function nextOrderId(slug: string): Promise<string> {
  const year = istanbulDateKey().slice(0, 4);
  const counterPath = `bayiler/${slug}/sayac-${year}.json`;
  const cur = await readJson<{ seq: number }>(counterPath);
  const seq = (Number(cur?.seq) || 0) + 1;
  const ok = await writeJson(counterPath, { seq });
  if (!ok) {
    const rand = Math.random().toString(36).slice(2, 6).toUpperCase();
    return `SIP-${year}-${rand}`;
  }
  return `SIP-${year}-${String(seq).padStart(3, "0")}`;
}

// ---- Müşteri takip linki ----
// Müşteri WhatsApp'tan aldığı linkle siparişinin durumunu görür. Link, gizli
// anahtarla imzalanır; tahmin edilemez, süresi yoktur.
export function trackingToken(slug: string, dateKey: string, orderId: string): string {
  return createHmac("sha256", authSecret())
    .update(`${slug}|${dateKey}|${orderId}`)
    .digest("hex")
    .slice(0, 20);
}

export function verifyTrackingToken(
  slug: string,
  dateKey: string,
  orderId: string,
  token: string
): boolean {
  return Boolean(token) && trackingToken(slug, dateKey, orderId) === token;
}

export function trackingPath(o: Pick<SavedOrder, "dealerSlug" | "dateKey" | "orderId">): string {
  const k = trackingToken(o.dealerSlug, o.dateKey, o.orderId);
  return `/takip?b=${o.dealerSlug}&d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}&k=${k}`;
}
