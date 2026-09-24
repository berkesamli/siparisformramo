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

export interface StokEslesme {
  item: StockItem;
  /** Yazılan kod stok koduyla birebir aynı (boşluk/tire/büyük-küçük harf farkı sayılmaz). */
  tam: boolean;
  /** Birebir değilse yön: yazılan kod adaylardan kısa (eksik) mı, uzun (fazla) mı. */
  yon?: "eksik" | "fazla";
  /** Olası kod sayısı: "KS4022-BİG" gibi eksik yazımda bütün BİG renkleri → belirsiz, miktar gösterilmez. */
  aday: number;
  /** Olası kodlar (en çok 6; arayüz "kodu tamamlayın" ipucunda listeler). */
  adaylar: string[];
}

interface StokKalem { item: StockItem; c: string }

// Normalize kodlar liste başına BİR kez hesaplanır (stokItems dizisi state'te sabit kalır; her tuş vuruşunda
// 700 kalemi yeniden normalize etmemek için). Aynı koda normalize olan kalemler (KS4022-BIG / KS4022-BİG —
// eski anlık görüntüler) burada birleştirilir: miktarlar toplanır, ilk görülen yazım gösterilir.
const kalemOnbellek = new WeakMap<StockItem[], StokKalem[]>();
function kalemler(items: StockItem[]): StokKalem[] {
  let h = kalemOnbellek.get(items);
  if (h) return h;
  const m = new Map<string, StokKalem>();
  for (const item of items) {
    const c = normalizeCode(item.code);
    if (!c) continue;
    const v = m.get(c);
    if (v) v.item = { ...v.item, ankaraMt: v.item.ankaraMt + item.ankaraMt, istanbulMt: v.item.istanbulMt + item.istanbulMt };
    else m.set(c, { item: { ...item }, c });
  }
  h = [...m.values()];
  kalemOnbellek.set(items, h);
  return h;
}

/**
 * Sipariş satırı için stok eşleşmesi. Bulanık benzerlik (harf hatası) hiç hesaplanmaz — yanlış modelin stoku
 * gösterilmesin (searchStock'un Levenshtein dalı burada zaten eşiğin altında kalıyordu).
 * - Birebir: tam=true.
 * - Eksik yazım (stok kodları yazılanla başlıyor; yoksa yazılanı içeriyor): her biri olası tamamlama → aday = hepsi;
 *   en kısası "en yakın" olarak döner. Arayüz aday > 1 ise miktar göstermez.
 * - Fazla yazım: yalnızca "yazılan = stok kodu + en çok 3 fazla karakter" (önek) sayılır; ortadan içerme
 *   ("GB185392OB" içindeki "5392" → 539-2 gibi) alakasız kısa kodu "en yakın" diye göstermez.
 * Hangi kodun eşleştiği item.code'da döner ki arayüz "stok bu koda ait" diye gösterebilsin.
 */
export function stokEslesme(items: StockItem[], code: string): StokEslesme | null {
  const q = normalizeCode(code);
  if (q.length < 2 || !items.length) return null;
  const hepsi = kalemler(items);
  const tam = hepsi.find((x) => x.c === q);
  if (tam) return { item: tam.item, tam: true, aday: 1, adaylar: [tam.item.code] };
  // Eksik yazım: önce önek, yoksa içerme
  let eksik = hepsi.filter((x) => x.c.startsWith(q));
  if (!eksik.length) eksik = hepsi.filter((x) => x.c.includes(q));
  if (eksik.length) {
    eksik.sort((a, b) => a.c.length - b.c.length);
    return { item: eksik[0].item, tam: false, yon: "eksik", aday: eksik.length, adaylar: eksik.slice(0, 6).map((x) => x.item.code) };
  }
  // Fazla yazım: en uzun (en özgül) kod en yakın; aynı uzunlukta olanlar aday
  const fazla = hepsi.filter((x) => q.startsWith(x.c) && x.c.length >= q.length - 3).sort((a, b) => b.c.length - a.c.length);
  if (!fazla.length) return null;
  const esit = fazla.filter((x) => x.c.length === fazla[0].c.length);
  return { item: esit[0].item, tam: false, yon: "fazla", aday: esit.length, adaylar: esit.slice(0, 6).map((x) => x.item.code) };
}
