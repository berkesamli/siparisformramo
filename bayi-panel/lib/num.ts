// Hesap hassasiyeti kuralları — fatura tutarlarının birebir tutması için.
//
// Kural: miktarlar (metre, m², adet) ve birim fiyatlar TAM hassasiyetle
// taşınır; yuvarlama yalnızca satır tutarında ve genel toplamlarda yapılır.
// Örn. 91,4 × 122 cm cam = 1,11508 m² (1,12 değil).

/** Kuruşa yuvarlar (2 hane) — satır tutarı ve toplamlar için. */
export const kurus = (n: number): number => Math.round((Number(n) || 0) * 100) / 100;

/**
 * Form kutusundaki sayıyı çözer — Türkçe ondalık virgülünü kabul eder:
 * "30,33" → 30.33. parseFloat virgülde kesip 30'a düşürüyordu; personel
 * fiyatı virgülle yazınca kuruşlar kayboluyordu.
 */
export const sayi = (s: string | number | undefined | null): number => {
  const v = parseFloat(String(s ?? "").trim().replace(",", "."));
  return Number.isFinite(v) ? v : 0;
};

/**
 * Hassasiyeti korur, yalnızca kayan nokta gürültüsünü temizler (6 hane).
 * Miktarlar ve birim fiyatlar için kullanılır.
 */
export const kesin = (n: number): number => Math.round((Number(n) || 0) * 1e6) / 1e6;

/** Miktar biçimi: gerektiği kadar ondalık, en fazla 5 hane. "1,11508" · "50" · "2,9" */
export const fmtQty = (n: number): string =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 5,
  });

/** Birim fiyat biçimi: en az 2, en fazla 4 hane. "71,25" · "105,925" */
export const fmtPrice = (n: number): string =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 4,
  });

/** Para biçimi: her zaman 2 hane. Tutarlar ve toplamlar için. */
export const fmtTL = (n: number): string =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
