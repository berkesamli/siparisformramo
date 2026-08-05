// Günlük stok Excel'ini (xls/xlsx) ayrıştırır — yalnızca sunucu tarafında kullanılır.
// Beklenen kolonlar: DEPO ADI, STOK İSMİ, MİKTAR (başlık satırı otomatik bulunur).
// Şimdilik yalnızca çerçeve profilleri ("PROFİL" içeren satırlar) alınır.

import * as XLSX from "xlsx";

export interface StockItem {
  code: string; // örn. GC065-1473BX
  ankaraMt: number;
  istanbulMt: number;
}

export interface StockData {
  updatedAt: string; // ISO
  sourceName: string;
  items: StockItem[];
}

const norm = (s: unknown) =>
  String(s ?? "")
    .toLocaleUpperCase("tr-TR")
    .replace(/İ/g, "I")
    .trim();

// STOK İSMİ içindeki ürün tanımı kelimeleri — kod bu kelimelerin
// öncesinde biter. Renk adları (KOREA SILVER, ANT GOLD, SHINING BLACK
// WITH YELLOW GOLD...) koda dahildir.
const TANIM_KELIMELERI = new Set([
  "LAMINE",
  "PROFIL",
  "SCAPPI",
  "PASPARTU",
  "KARTON",
  "TELALI",
  "KADIFE",
  "DÜZ",
]);

// "KS3514-KOREA SILVER LAMİNE PROFİL" → "KS3514-KOREA SILVER"
export function extractCode(name: string): string {
  const tokens = String(name).trim().split(/\s+/);
  const kept: string[] = [];
  for (const t of tokens) {
    if (TANIM_KELIMELERI.has(norm(t))) break;
    kept.push(t.toLocaleUpperCase("tr-TR"));
  }
  return kept.join(" ");
}

export function parseStockWorkbook(buffer: Buffer, sourceName: string): StockData {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  // Başlık satırını bul
  let headerIdx = -1;
  let colDepo = -1;
  let colName = -1;
  let colQty = -1;
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const cells = rows[i].map(norm);
    const d = cells.findIndex((c) => c.includes("DEPO ADI"));
    const n = cells.findIndex((c) => c.includes("STOK ISMI"));
    const q = cells.findIndex((c) => c === "MIKTAR" || c.includes("MIKTAR"));
    if (d >= 0 && n >= 0 && q >= 0) {
      headerIdx = i;
      colDepo = d;
      colName = n;
      colQty = q;
      break;
    }
  }
  if (headerIdx < 0) {
    throw new Error(
      "Excel'de 'DEPO ADI', 'STOK İSMİ' ve 'MİKTAR' başlıkları bulunamadı. Dosya formatını kontrol edin."
    );
  }

  const map = new Map<string, StockItem>();
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const depot = norm(row[colDepo]);
    const name = String(row[colName] ?? "").trim();
    const qty = parseFloat(String(row[colQty]).replace(",", "."));
    if (!name || !depot || !isFinite(qty) || qty <= 0) continue;
    // Şimdilik yalnızca çerçeve profilleri
    if (!norm(name).includes("PROFIL")) continue;

    // "GC065-1473BX LAMİNE PROFİL"            → kod = GC065-1473BX
    // "KS3514-KOREA SILVER LAMİNE PROFİL"     → kod = KS3514-KOREA SILVER
    // Renk adı birden çok kelime olabilir; kod, açıklama kelimesine
    // (LAMİNE/PROFİL vb.) kadar olan kısımdır — ilk boşlukta kesilmez.
    const code = extractCode(name);
    if (!code) continue;

    let item = map.get(code);
    if (!item) {
      item = { code, ankaraMt: 0, istanbulMt: 0 };
      map.set(code, item);
    }
    if (depot.includes("ANKARA")) item.ankaraMt += qty;
    else if (depot.includes("ISTANBUL")) item.istanbulMt += qty;
  }

  const items = Array.from(map.values())
    .map((it) => ({
      code: it.code,
      ankaraMt: Math.round(it.ankaraMt * 100) / 100,
      istanbulMt: Math.round(it.istanbulMt * 100) / 100,
    }))
    .sort((a, b) => a.code.localeCompare(b.code, "tr"));

  if (!items.length) {
    throw new Error("Excel'de profil stoğu bulunamadı (PROFİL içeren satır yok).");
  }

  return {
    updatedAt: new Date().toISOString(),
    sourceName,
    items,
  };
}
