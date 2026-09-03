// Bayi fiyatlandırma modeli.
//
// Çerçeve:   toptan liste fiyatı (USD/mt, data/catalog.ts) × bayi çarpanı × USD kuru = TL/m
// Paspartu:  bayi başına TL/m² (tür kodu sabit, fiyat bayinin)
// Cam:       bayi başına TL/m²
// Baskı:     bayi başına USD/m² (TL = USD × kur)
// İşçilik:   kalem başına sabit TL (opsiyonel)
//
// Aşağıdaki varsayılanlar Olga'nın perakende mağazasında kullandığı değerlerdir;
// yeni bayi bu değerlerle başlar, panelden kendine göre değiştirir.

export interface MatType {
  code: string; // sabit — renk paleti ve üretim PDF'i bu koda bağlı
  name: string;
  price: number; // TL/m²
  icon?: string;
}

export interface GlassType {
  name: string;
  price: number; // TL/m²
  desc: string;
  icon: string;
}

export interface PrintType {
  name: string;
  usdPerM2: number;
  desc: string;
  icon: string;
}

export interface DealerPricing {
  frameFactor: number; // toptan USD liste fiyatı çarpanı
  usdRateMode: "auto" | "manual";
  usdRate: number; // manuel kur (auto modda yedek)
  mats: MatType[];
  glasses: GlassType[];
  prints: PrintType[];
  laborTL: number; // kalem başına işçilik (0 = yok)
  updatedAt?: string;
}

export const DEFAULT_FRAME_FACTOR = 5;

export const DEFAULT_MATS: MatType[] = [
  { code: "-", name: "Paspartu Yok", price: 0, icon: "🚫" },
  { code: "DK", name: "Düz Karton", price: 1500, icon: "🟨" },
  { code: "PK", name: "Pamuk Karton", price: 2500, icon: "🪵" },
  { code: "AG", name: "Altın-Gümüş", price: 3000, icon: "✨" },
  { code: "KDF", name: "Kadife", price: 5500, icon: "🧵" },
  { code: "PR", name: "Premium", price: 5500, icon: "🟪" },
];

export const DEFAULT_GLASSES: GlassType[] = [
  { name: "Cam Yok", price: 0, desc: "Camsız teslim", icon: "🚫" },
  { name: "Düz Cam", price: 2000, desc: "Standart şeffaf cam", icon: "🪟" },
  { name: "Mat Cam", price: 2000, desc: "Yansıma yapmayan mat cam", icon: "🌫️" },
  { name: "PVC Cam", price: 2000, desc: "Kırılmaz hafif PVC (pleksi)", icon: "🛡️" },
  { name: "Müze Camı", price: 12000, desc: "UV korumalı premium cam", icon: "🏛️" },
];

export const DEFAULT_PRINTS: PrintType[] = [
  { name: "Baskı Yok", usdPerM2: 0, desc: "Baskı istemiyorum", icon: "🚫" },
  { name: "Polyester Baskı", usdPerM2: 52.5, desc: "Canlı renkler, ekonomik", icon: "🖼️" },
  { name: "Deri Bez Baskı", usdPerM2: 56.7, desc: "Deri dokulu özel bez", icon: "🟤" },
  { name: "Pamuk Bez Baskı", usdPerM2: 60.9, desc: "Doğal pamuk kanvas", icon: "🧶" },
  { name: "HP Mat Fine Art Baskı", usdPerM2: 69.3, desc: "Müze kalitesi fine art", icon: "🎨" },
];

export function defaultPricing(): DealerPricing {
  return {
    frameFactor: DEFAULT_FRAME_FACTOR,
    usdRateMode: "auto",
    usdRate: 0,
    mats: DEFAULT_MATS.map((m) => ({ ...m })),
    glasses: DEFAULT_GLASSES.map((g) => ({ ...g })),
    prints: DEFAULT_PRINTS.map((p) => ({ ...p })),
    laborTL: 0,
  };
}

/**
 * Bayiden gelen ayarları varsayılan listeyle birleştirir: kodlar/isimler
 * sabittir (üretim ve önizleme onlara bağlı), yalnızca fiyatlar bayinindir.
 * Böylece ileride yeni bir cam/paspartu türü eklenince eski bayilerde de görünür.
 */
export function normalizePricing(raw: unknown): DealerPricing {
  const d = defaultPricing();
  const r = (raw && typeof raw === "object" ? raw : {}) as Record<string, any>;
  const num = (v: unknown, fallback: number, min = 0) => {
    const n = Number(v);
    return Number.isFinite(n) && n >= min ? n : fallback;
  };
  const priceOf = (list: unknown, key: "code" | "name", id: string): number | undefined => {
    if (!Array.isArray(list)) return undefined;
    const hit = list.find((x: any) => x && x[key] === id);
    if (!hit) return undefined;
    const n = Number(hit.price ?? hit.usdPerM2);
    return Number.isFinite(n) && n >= 0 ? n : undefined;
  };
  return {
    frameFactor: num(r.frameFactor, d.frameFactor, 0.01),
    usdRateMode: r.usdRateMode === "manual" ? "manual" : "auto",
    usdRate: num(r.usdRate, 0),
    mats: d.mats.map((m) =>
      m.code === "-" ? m : { ...m, price: priceOf(r.mats, "code", m.code) ?? m.price }
    ),
    glasses: d.glasses.map((g) =>
      g.name === "Cam Yok" ? g : { ...g, price: priceOf(r.glasses, "name", g.name) ?? g.price }
    ),
    prints: d.prints.map((p) =>
      p.name === "Baskı Yok"
        ? p
        : { ...p, usdPerM2: priceOf(r.prints, "name", p.name) ?? p.usdPerM2 }
    ),
    laborTL: num(r.laborTL, 0),
    updatedAt: typeof r.updatedAt === "string" ? r.updatedAt : undefined,
  };
}

export const ORDER_STATUSES = [
  "Beklemede",
  "Hazırlanıyor",
  "Hazır",
  "Teslim Edildi",
  "İptal",
] as const;
export type OrderStatus = (typeof ORDER_STATUSES)[number];

export type PaymentStatus = "bekliyor" | "kismi" | "odendi";
export const PAYMENT_LABELS: Record<PaymentStatus, string> = {
  bekliyor: "Ödeme Bekliyor",
  kismi: "Kısmi Ödendi",
  odendi: "Ödendi",
};

export function toMM(value: number, unit: "cm" | "mm"): number {
  return unit === "cm" ? value * 10 : value;
}

export interface CostInput {
  wMM: number;
  hMM: number;
  matTop: number;
  matRight: number;
  matBottom: number;
  matLeft: number;
  framePriceTL: number; // TL/m
  matPrice: number; // dış paspartu TL/m² (0 = yok)
  hasMat: boolean;
  doubleMat: boolean;
  innerMatPrice: number;
  zeminEnabled: boolean;
  zeminPrice: number;
  glassPrice: number; // TL/m²
  printUsdPerM2: number;
  usdRate: number;
  laborTL: number;
}

export interface Costs {
  frameCost: number;
  matCost: number;
  glassCost: number;
  printCost: number;
  laborCost: number;
  itemTotal: number;
}

// Hesap kuralı (Olga perakende ile aynı): dış ölçü = eser + paspartu kenarları;
// çevre = 2×(en+boy) + 0,30 m fire; cam/paspartu alan üzerinden, baskı eserin
// kendi alanı üzerinden.
export function computeCosts(inp: CostInput): Costs {
  const tw = inp.wMM + inp.matLeft + inp.matRight;
  const th = inp.hMM + inp.matTop + inp.matBottom;
  const area = (tw / 1000) * (th / 1000);
  const perim = (2 * (tw + th)) / 1000 + 0.3;

  const frameCost = perim * inp.framePriceTL;

  let matCost = 0;
  if (inp.hasMat) {
    matCost += area * inp.matPrice;
    if (inp.doubleMat) matCost += area * inp.innerMatPrice;
    if (inp.zeminEnabled) matCost += area * inp.zeminPrice;
  }

  const glassCost = area * inp.glassPrice;
  const printArea = (inp.wMM / 1000) * (inp.hMM / 1000);
  const printCost =
    inp.printUsdPerM2 > 0 ? printArea * inp.printUsdPerM2 * inp.usdRate : 0;
  const laborCost = inp.laborTL > 0 ? inp.laborTL : 0;

  const itemTotal = frameCost + matCost + glassCost + printCost + laborCost;
  return { frameCost, matCost, glassCost, printCost, laborCost, itemTotal };
}
