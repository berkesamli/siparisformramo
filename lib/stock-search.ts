// Stok kodu bulanık arama (client-safe).
// Kullanıcılar "gc065-1473" gibi kısmi/eksik yazar; GC065-1473BX gibi
// varyantları da yakalamak için normalize + benzerlik eşleşmesi yapılır.

import type { StockItem } from "./stock-parse";

export const BOY_LENGTH = 2.9; // metre → boy çevirimi

export function toBoy(mt: number): number {
  return Math.floor(mt / BOY_LENGTH);
}

export function normalizeCode(s: string): string {
  return s
    .toLocaleUpperCase("tr-TR")
    .replace(/İ/g, "I")
    .replace(/[^A-Z0-9]/g, "");
}

function levenshtein(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  if (!m) return n;
  if (!n) return m;
  let prev = new Array(n + 1);
  let curr = new Array(n + 1);
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[n];
}

function similarity(a: string, b: string): number {
  const maxLen = Math.max(a.length, b.length);
  if (!maxLen) return 1;
  return 1 - levenshtein(a, b) / maxLen;
}

export interface StockMatch {
  item: StockItem;
  score: number; // 1 = birebir
}

/**
 * Sorguya göre stok kalemlerini puanlayıp sıralar.
 * - Birebir eşleşme: 1.0
 * - Kod, sorgu ile başlıyorsa (GC0651473 → GC0651473BX): 0.95
 * - Kod sorguyu içeriyorsa / sorgu kodu içeriyorsa: 0.9
 * - Benzerlik oranı ≥ minScore olanlar (harf hatalarını tolere eder)
 */
export function searchStock(
  items: StockItem[],
  query: string,
  minScore = 0.72,
  limit = 30
): StockMatch[] {
  const q = normalizeCode(query);
  if (q.length < 2) return [];

  const matches: StockMatch[] = [];
  for (const item of items) {
    const c = normalizeCode(item.code);
    let score = 0;
    if (c === q) score = 1;
    else if (c.startsWith(q) || q.startsWith(c)) score = 0.95;
    else if (c.includes(q) || q.includes(c)) score = 0.9;
    else {
      // Uzunluklar çok farklıysa kısmi karşılaştır (ön ek üzerinden)
      const cut = Math.min(c.length, Math.max(q.length, 4) + 2);
      const s = Math.max(similarity(c, q), similarity(c.slice(0, cut), q));
      if (s >= minScore) score = s * 0.85; // tam eşleşmelerin altında kalsın
    }
    if (score > 0) matches.push({ item, score });
  }

  return matches.sort((a, b) => b.score - a.score).slice(0, limit);
}
