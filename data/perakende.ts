// Olga Çerçeve — Perakende çerçeveletme (online çerçeve) verileri.
// Buradaki tüm fiyatlar müşteriye gösterilen PERAKENDE satış fiyatlarıdır.

// [kod, hex, metalik?]
export type PaletteColor = [string, string, boolean?];

// Paspartu renk paletleri — anahtar: paspartu türünün m² fiyatı (TL)
export const PASPARTU_COLORS: Record<number, PaletteColor[]> = {
  1250: [
    ["W107", "#1b227c"], ["W108", "#ffa711"], ["W109", "#ff7d01"], ["W110", "#ff7d01"],
    ["W111", "#8468b3"], ["W112", "#9e508c"], ["W125", "#f23a04"], ["W131", "#fd0b35"],
    ["W132", "#871f26"], ["W135", "#847c89"], ["W136", "#fbef8b"], ["W140", "#f3f4a8"],
    ["W141", "#ebe5cb"], ["W142", "#fbf27b"], ["W146", "#f4efe9"], ["W150", "#fafbeb"],
    ["W151", "#eeede9"], ["W152", "#ecf1ed"], ["W153", "#ebdf95"], ["W154", "#ca917d"],
    ["W155", "#f5ecda"], ["W156", "#e2d2b9"], ["W158", "#e7c983"], ["W159", "#c0b7b0"],
    ["W160", "#d9d4c0"], ["W161", "#e5e8a1"], ["W162", "#98b795"], ["W163", "#48813c"],
    ["W164", "#587f50"], ["W165", "#024232"], ["W166", "#c59b5f"], ["W167", "#76391d"],
    ["W168", "#361b12"], ["W170", "#598180"], ["W171", "#001825"], ["W172", "#000000"],
    ["W176", "#cfb793"], ["W177", "#c99f6d"], ["W178", "#d7cdaa"], ["W181", "#ad2029"],
    ["W182", "#431a1e"], ["W187", "#e7af8c"], ["W188", "#7f853d"], ["W189", "#451b0b"],
    ["W190", "#ffc93e"], ["W191", "#005892"], ["W192", "#d83e08"], ["W193", "#3f5855"],
    ["W194", "#72b7be"], ["W195", "#71939c"], ["W196", "#012f49"], ["W197", "#024c33"],
    ["W198", "#2b1c17"], ["W199", "#c2400e"], ["W253", "#f3da98"], ["W406", "#3d3832"],
  ],
  1350: [
    ["G163", "#2a4018"], ["G167", "#7e5121"], ["G170", "#495773"], ["G175", "#dec49c"],
    ["G182", "#331f21"], ["G188", "#858060"], ["G189", "#4c2614"], ["G197", "#192f22"],
    ["W270", "#e9dbca"], ["W274", "#dcb4b0"], ["W277", "#bdc5b5"], ["W279", "#b1bcbd"],
    ["W284", "#f7dfc7"], ["W289", "#ded1ad"], ["W290", "#d4c5a4"],
  ],
  1700: [
    ["W232", "#d4af37", true], ["W233", "#c0c0c0", true],
  ],
  1800: [
    ["700", "#c5c1bf"], ["708", "#385641"], ["711", "#792c27"], ["718", "#294a6f"],
    ["719", "#2e3f46"], ["720", "#30322f"], ["721", "#b2a287"], ["722", "#c1b396"],
    ["731", "#867754"], ["W600", "#bbb5ab"], ["W601", "#b6a68c"], ["W602", "#b7af99"],
    ["W603", "#c8bda8"], ["W604", "#c0ae8f"], ["W605", "#a38860"], ["W606", "#8c6b41"],
    ["W607", "#968863"], ["W608", "#a1a598"], ["W610", "#bda170"], ["W611", "#d4b066"],
    ["W612", "#83857e"], ["W613", "#885127"], ["W614", "#2d302d"],
  ],
  3000: [
    ["750", "#f6d8bc"], ["752", "#d4ad90"], ["753", "#8e4a01"], ["754", "#7a2a05"],
    ["755", "#34130a"], ["756", "#cb8f49"], ["757", "#9a9e7d"], ["758", "#00340d"],
    ["760", "#ad6c70"], ["761", "#b00305"], ["762", "#8f030c"], ["763", "#570e1f"],
    ["765", "#1e3a45"], ["767", "#050d32"], ["770", "#a49d97"], ["772", "#19191b"],
    ["777", "#6c6d33"], ["778", "#c9c8b4"], ["779", "#0e0e16"], ["780", "#6f3645"],
    ["781", "#94063e"], ["782", "#471546"], ["783", "#844921"], ["786", "#9b8b7b"],
    ["787", "#533f5a"],
  ],
};

export interface MatType {
  code: string;
  name: string;
  price: number; // TL/m² perakende
  icon?: string;
}

export const MAT_TYPES: MatType[] = [
  { code: "-", name: "Paspartu Yok", price: 0, icon: "🚫" },
  { code: "DK", name: "Düz Karton", price: 1250, icon: "🟨" },
  { code: "DMK", name: "Damarlı Karton", price: 1350, icon: "🪵" },
  { code: "AG", name: "Altın-Gümüş", price: 1700, icon: "✨" },
  { code: "DOK", name: "Dokulu", price: 1800, icon: "🧵" },
  { code: "KDF", name: "Kadife", price: 3000, icon: "🟪" },
];

// İç paspartu / zemin seçenekleri ("Paspartu Yok" hariç)
export const INNER_MAT_TYPES: MatType[] = MAT_TYPES.filter((m) => m.price > 0);

export interface GlassType {
  name: string;
  price: number; // TL/m² perakende
  desc: string;
  icon: string;
}

export const GLASS_TYPES: GlassType[] = [
  { name: "Cam Yok", price: 0, desc: "Camsız teslim", icon: "🚫" },
  { name: "Düz Cam", price: 1250, desc: "Standart şeffaf cam", icon: "🪟" },
  { name: "Mat Cam", price: 1750, desc: "Yansıma yapmayan mat cam", icon: "🌫️" },
  { name: "PVC Cam", price: 1250, desc: "Kırılmaz hafif PVC", icon: "🛡️" },
  { name: "Müze Camı", price: 12000, desc: "UV korumalı premium cam", icon: "🏛️" },
];

export interface PrintType {
  name: string;
  usdPerM2: number; // USD/m² perakende satış fiyatı (KDV dahil) — TL = usdPerM2 × kur
  desc: string;
  icon: string;
}

export const PRINT_TYPES: PrintType[] = [
  { name: "Baskı Yok", usdPerM2: 0, desc: "Baskı istemiyorum", icon: "🚫" },
  { name: "Polyester Baskı", usdPerM2: 52.5, desc: "Canlı renkler, ekonomik", icon: "🖼️" },
  { name: "Deri Bez Baskı", usdPerM2: 56.7, desc: "Deri dokulu özel bez", icon: "🟤" },
  { name: "Pamuk Bez Baskı", usdPerM2: 60.9, desc: "Doğal pamuk kanvas", icon: "🧶" },
  { name: "HP Mat Fine Art Baskı", usdPerM2: 69.3, desc: "Müze kalitesi fine art", icon: "🎨" },
];

export const RETAIL_STATUSES = [
  "Beklemede",
  "Hazırlanıyor",
  "Hazır",
  "Teslim Edildi",
  "İptal",
] as const;

export type RetailStatus = (typeof RETAIL_STATUSES)[number];

export function toMM(value: number, unit: "cm" | "mm"): number {
  return unit === "cm" ? value * 10 : value;
}

export interface RetailCostInput {
  wMM: number; // eser genişliği (mm)
  hMM: number; // eser yüksekliği (mm)
  matTop: number; // paspartu kenarları (mm)
  matRight: number;
  matBottom: number;
  matLeft: number;
  framePriceTL: number; // TL/metre
  matPrice: number; // dış paspartu TL/m² (0 = yok)
  doubleMat: boolean;
  innerMatPrice: number;
  zeminEnabled: boolean;
  zeminPrice: number;
  glassPrice: number; // TL/m²
  printUsdPerM2: number; // USD/m²
  usdRate: number;
}

export interface RetailCosts {
  frameCost: number;
  matCost: number;
  glassCost: number;
  printCost: number;
  itemTotal: number;
}

// Orijinal hesap: dış ölçü = eser + paspartu kenarları;
// çevre = 2×(en+boy) + 0.30 m fire; cam/paspartu alan üzerinden,
// baskı eserin kendi alanı üzerinden hesaplanır.
export function computeRetailCosts(inp: RetailCostInput): RetailCosts {
  const tw = inp.wMM + inp.matLeft + inp.matRight;
  const th = inp.hMM + inp.matTop + inp.matBottom;
  const area = (tw / 1000) * (th / 1000);
  const perim = (2 * (tw + th)) / 1000 + 0.3;

  const frameCost = perim * inp.framePriceTL;

  let matCost = 0;
  if (inp.matPrice > 0) {
    matCost += area * inp.matPrice;
    if (inp.doubleMat) matCost += area * inp.innerMatPrice;
    if (inp.zeminEnabled) matCost += area * inp.zeminPrice;
  }

  const glassCost = area * inp.glassPrice;

  const printArea = (inp.wMM / 1000) * (inp.hMM / 1000);
  const printCost =
    inp.printUsdPerM2 > 0 ? printArea * inp.printUsdPerM2 * inp.usdRate : 0;

  const itemTotal = frameCost + matCost + glassCost + printCost;
  return { frameCost, matCost, glassCost, printCost, itemTotal };
}
