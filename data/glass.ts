// Cam ve ayna plaka ölçüleri — Form.html'den taşındı.

export interface PlateSize {
  label: string;
  en: number; // cm
  boy: number; // cm
  mm?: number;
}

export const GLASS_TYPES: { key: string; name: string }[] = [
  { key: "duz", name: "Düz Cam" },
  { key: "mat", name: "Mat Cam" },
  { key: "muze", name: "Müze Camı" },
];

export const GLASS_SIZES: Record<string, PlateSize[]> = {
  mat: [{ label: "122 × 183 cm", en: 122, boy: 183 }],
  duz: [
    { label: "122 × 183 cm", en: 122, boy: 183 },
    { label: "121 × 161 cm", en: 121, boy: 161 },
    { label: "91,4 × 122 cm", en: 91.4, boy: 122 },
  ],
  muze: [{ label: "122 × 172,5 cm", en: 122, boy: 172.5 }],
};

export const AYNA_SIZES: PlateSize[] = [
  { label: "91,4 × 122 cm — 2mm", en: 91.4, boy: 122, mm: 2 },
  { label: "122 × 183 cm — 3mm", en: 122, boy: 183, mm: 3 },
  { label: "122 × 183 cm — 4mm", en: 122, boy: 183, mm: 4 },
];

// Plaka alanı tam hassasiyetle döner — yuvarlanmaz.
// 91,4 × 122 cm → 1,11508 m² (1,12 değil). Fatura tutarlarının tutması için
// yuvarlama yalnızca satır tutarında yapılır.
export function plateM2(s: PlateSize): number {
  return Math.round((s.en / 100) * (s.boy / 100) * 1e6) / 1e6;
}
