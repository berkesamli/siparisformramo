// Güncel stok verisine tek noktadan erişim.
// Önce Vercel Blob'daki günlük yüklenen dosya, yoksa repo içindeki snapshot.

import type { StockData } from "./stock-parse";
import snapshot from "@/data/stock-snapshot.json";

export const STOCK_BLOB_PATH = "stock/latest.json";

export function stockBlobConfigured(): boolean {
  return Boolean(process.env.BLOB_STORE_ID || process.env.BLOB_READ_WRITE_TOKEN);
}

async function readFromBlob(): Promise<StockData | null> {
  if (!stockBlobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const result = await get(STOCK_BLOB_PATH, { access: "private", useCache: false });
    if (!result || result.statusCode !== 200 || !result.stream) return null;
    const text = await new Response(result.stream).text();
    return JSON.parse(text) as StockData;
  } catch (err) {
    console.error("Blob okunamadı:", err);
    return null;
  }
}

export async function getStockData(): Promise<StockData> {
  return (await readFromBlob()) ?? (snapshot as StockData);
}
