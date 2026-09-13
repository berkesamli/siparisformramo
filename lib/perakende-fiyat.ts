// =========================================================
// PERAKENDE FİYAT ÇEKİRDEĞİ — olgacerceve.com hesaplayıcısından taşındı
// (websitesi-er-evehesaplay-c- deposu, server/lib/fiyat.js v1.2.0).
// Sihirbaz (RetailWizard) ve sunucu (POST /api/perakende/orders) fiyatı
// AYNI koddan hesaplar; sunucu istemcinin gönderdiği tutarı yeniden
// hesaplayıp doğrular — websitedeki çözülmüş kuralların birebir karşılığı.
//
// Kurallar (websiteyle aynı):
// - Çerçeve = (çevre + 0,30 m fire) × profil metre fiyatı
// - Paspartu = iç alan (m²) × karton m² fiyatı; çift paspartuda ikinci
//   karton da aynı alandan; iç şerit (altMontaj) iç ölçüye 2× eklenir
// - Pencere kesim ücreti: ilk pencere hariç, pencere başına (karton türüne
//   göre); çift paspartuda iki katman
// - Cam = iç alan (m²) × cam m² fiyatı; kasa (kanvas) çerçevede cam ve
//   paspartu yok, 5 mm gölge boşluğu iç ölçüye eklenir
// - İç ölçü = eser (veya pencereli alan) + paspartu kenarları
//   + 2 × iç şerit (çift paspartu) + 2 × kasa gölge boşluğu
// Panele özgü eklentiler: zemin kartonu (3. katman, aynı alandan) ve
// baskı (eser alanı × USD/m² × kur) — websitede yoktur, aynen korunur.
// =========================================================

export const PERAKENDE_SABIT = {
  FIRE_M: 0.3, // her çerçevede eklenen fire (m)
  MAT_APERTURE_OVERLAP_MM: 5, // açıklık eserden her kenarda bu kadar dar kesilir
  CANVAS_SHADOW_GAP_MM: 5, // kasa (kanvas) çerçevede tuval-kasa boşluğu (her kenar)
  MIN_MAT_BRIDGE_MM: 20, // pencereler arası en az köprü
  MIN_MAT_EDGE_MM: 20, // paspartu kenarı en az
  MAX_MAT_EDGE_MM: 300, // paspartu kenarı en fazla
  MAT_SHEET_SHORT_MM: 800, // en büyük paspartu tabakası 80 × 120 cm
  MAT_SHEET_LONG_MM: 1200,
  GLASS_MAX_SHORT_MM: 1000, // en büyük cam plakası 100 × 140 cm
  GLASS_MAX_LONG_MM: 1400,
  MAX_OUTER_MM: 2900, // profil boyu 290 cm: iç/dış ölçü bunu aşamaz
  MIN_ART_MM: 15, // eser kenarı en az
  MAX_WINDOWS: 9,
  MAX_MOUNTING_MM: 50, // iç şerit üst sınırı
} as const;

// İlk pencere hariç, pencere başına kesim ücreti (TL) — anahtar: karton m²
// fiyatı. Websitedeki tablo + panele özgü Pamuk Karton (2500).
export const PENCERE_UCRETI: Record<number, number> = {
  1500: 40,
  2500: 50,
  3000: 50,
  5500: 75,
};

export interface PencereDuzen {
  id: string; // "2x3"
  rows: number;
  cols: number;
  label: string;
}

/** n pencere için olası satır×sütun düzenleri (websitedeki liste). */
export function pencereDuzenleri(n: number): PencereDuzen[] {
  const out: PencereDuzen[] = [];
  if (!(n >= 2)) return out;
  out.push({ id: `1x${n}`, rows: 1, cols: n, label: "Yan yana" });
  out.push({ id: `${n}x1`, rows: n, cols: 1, label: "Alt alta" });
  for (let r = 2; r <= Math.floor(n / 2); r++) {
    if (n % r === 0) {
      const c = n / r;
      if (c >= 2)
        out.push({ id: `${r}x${c}`, rows: r, cols: c, label: `${r} satır × ${c} sütun` });
    }
  }
  return out;
}

export function varsayilanDuzen(n: number): PencereDuzen | null {
  const pref: Record<number, string> = { 4: "2x2", 6: "2x3", 8: "2x4", 9: "3x3" };
  const list = pencereDuzenleri(n);
  return list.find((a) => a.id === pref[n]) || list[0] || null;
}

/** Fotoğraf ölçüsüne göre önerilen pencere aralığı (mm). */
export function varsayilanAralik(wMM: number, hMM: number): number {
  const uzun = Math.max(wMM || 0, hMM || 0);
  return uzun < 200 ? 20 : uzun < 297 ? 25 : 30;
}

export interface PencereYerlesim {
  rows: number;
  cols: number;
  count: number;
  gap: number;
  mounting: number;
  apW: number; // iç açıklık (fotoğrafı tutan) — fotoğraf her kenarda 5 mm altta kalır
  apH: number;
  fieldW: number; // tüm fotoğrafları kapsayan alan — kenarlar bundan ölçülür
  fieldH: number;
}

/**
 * Pencereli paspartu geometrisi (websitedeki computeWindowLayout):
 * ölçü girişi TEK fotoğrafın ölçüsüdür; alan (field) eser ölçüsünün yerine
 * geçer. mounting: çift paspartuda pencere çevresindeki iç şerit.
 */
export function pencereYerlesim(
  photoW: number,
  photoH: number,
  rows: number,
  cols: number,
  gap: number,
  mounting: number
): PencereYerlesim {
  const ov = PERAKENDE_SABIT.MAT_APERTURE_OVERLAP_MM;
  const m = Math.max(0, mounting || 0);
  const apW = Math.max(10, photoW - 2 * ov);
  const apH = Math.max(10, photoH - 2 * ov);
  const pitchW = apW + 2 * m + gap;
  const pitchH = apH + 2 * m + gap;
  const fieldW = photoW + (cols - 1) * pitchW;
  const fieldH = photoH + (rows - 1) * pitchH;
  return { rows, cols, count: rows * cols, gap, mounting: m, apW, apH, fieldW, fieldH };
}

export interface PerakendeGirdi {
  wMM: number; // tek fotoğraf/eser ölçüsü (mm)
  hMM: number;
  kenar: { ust: number; alt: number; sol: number; sag: number };
  matPrice: number; // 0 = paspartu yok (TL/m²)
  doubleMat: boolean;
  innerMatPrice: number;
  icSeritMm: number; // çift paspartuda pencere çevresindeki iç şerit (altMontaj)
  zeminEnabled: boolean;
  zeminPrice: number;
  pencereSayisi: number; // 1..9 (yalnız paspartu varken >1 olabilir)
  pencereRows?: number;
  pencereCols?: number;
  pencereAralikMm?: number;
  camPrice: number; // TL/m² (0 = cam yok)
  camMaxKisaMM?: number; // cam türünün üretim sınırı (engel)
  camMaxUzunMM?: number;
  camUyariKisaMM?: number; // aşılırsa uyarı (2 mm gerçek camda kırılma riski)
  camUyariUzunMM?: number;
  kasa: boolean; // kanvas/kasa çerçeve — cam ve paspartu uygulanmaz
  framePriceTL: number; // TL/metre
  printUsdPerM2: number;
  usdRate: number;
}

function num(v: unknown, d = 0): number {
  const n =
    typeof v === "string" ? parseFloat(v.replace(",", ".")) : Number(v);
  return Number.isFinite(n) ? n : d;
}

interface NormalGirdi {
  wMM: number;
  hMM: number;
  kenar: { ust: number; alt: number; sol: number; sag: number };
  matPrice: number;
  innerMatPrice: number;
  icSeritMm: number;
  zeminPrice: number;
  pencereSayisi: number;
  rows: number;
  cols: number;
  aralik: number;
  camPrice: number;
  kasa: boolean;
  framePriceTL: number;
  printUsdPerM2: number;
  usdRate: number;
}

function normalize(g: PerakendeGirdi): NormalGirdi {
  const kasa = !!g.kasa;
  const matPrice = kasa ? 0 : Math.max(0, num(g.matPrice));
  const innerMatPrice =
    matPrice > 0 && g.doubleMat ? Math.max(0, num(g.innerMatPrice)) : 0;
  const zeminPrice =
    matPrice > 0 && g.zeminEnabled ? Math.max(0, num(g.zeminPrice)) : 0;
  const pencereSayisi =
    matPrice > 0 ? Math.max(1, Math.round(num(g.pencereSayisi, 1))) : 1;
  const duzen =
    pencereSayisi > 1
      ? (g.pencereRows && g.pencereCols
          ? { rows: Math.round(num(g.pencereRows)), cols: Math.round(num(g.pencereCols)) }
          : varsayilanDuzen(pencereSayisi)) || { rows: 1, cols: pencereSayisi }
      : { rows: 1, cols: 1 };
  return {
    wMM: num(g.wMM),
    hMM: num(g.hMM),
    kenar: {
      ust: num(g.kenar?.ust),
      alt: num(g.kenar?.alt),
      sol: num(g.kenar?.sol),
      sag: num(g.kenar?.sag),
    },
    matPrice,
    innerMatPrice,
    icSeritMm: innerMatPrice > 0 ? num(g.icSeritMm, 5) : 0,
    zeminPrice,
    pencereSayisi,
    rows: Math.max(1, duzen.rows),
    cols: Math.max(1, duzen.cols),
    aralik:
      pencereSayisi > 1
        ? Math.max(
            PERAKENDE_SABIT.MIN_MAT_BRIDGE_MM,
            num(g.pencereAralikMm, varsayilanAralik(num(g.wMM), num(g.hMM)))
          )
        : 0,
    camPrice: kasa ? 0 : Math.max(0, num(g.camPrice)),
    kasa,
    framePriceTL: Math.max(0, num(g.framePriceTL)),
    printUsdPerM2: Math.max(0, num(g.printUsdPerM2)),
    usdRate: Math.max(0, num(g.usdRate)),
  };
}

/** Pencereli düzende eserin yerine geçen alan (field) ölçüsü. */
function alanOlcusu(n: NormalGirdi): { en: number; boy: number; yerlesim: PencereYerlesim | null } {
  if (n.pencereSayisi > 1) {
    const y = pencereYerlesim(n.wMM, n.hMM, n.rows, n.cols, n.aralik, n.icSeritMm);
    return { en: y.fieldW, boy: y.fieldH, yerlesim: y };
  }
  return { en: n.wMM, boy: n.hMM, yerlesim: null };
}

export interface PerakendeSonuc {
  icEn: number; // çerçeve iç ölçüsü (= paspartu dış ölçüsü / kasa iç açıklığı)
  icBoy: number;
  alanM2: number;
  metreToplam: number; // çevre + fire
  frameCost: number;
  matCost: number; // paspartu katmanları + zemin + pencere kesim ücreti
  pencereKesim: number; // matCost içindeki kesim ücreti payı
  glassCost: number;
  printCost: number;
  itemTotal: number;
  yerlesim: PencereYerlesim | null;
}

export function hesaplaPerakende(g: PerakendeGirdi): PerakendeSonuc {
  const n = normalize(g);
  const alan = alanOlcusu(n);
  const hasMat = n.matPrice > 0;
  const kenar = hasMat ? n.kenar : { ust: 0, alt: 0, sol: 0, sag: 0 };
  const serit = n.innerMatPrice > 0 ? n.icSeritMm * 2 : 0;
  const kasaPay = n.kasa ? PERAKENDE_SABIT.CANVAS_SHADOW_GAP_MM * 2 : 0;
  const icEn = alan.en + kenar.sol + kenar.sag + serit + kasaPay;
  const icBoy = alan.boy + kenar.ust + kenar.alt + serit + kasaPay;
  const alanM2 = (icEn / 1000) * (icBoy / 1000);
  const metreToplam = (2 * (icEn + icBoy)) / 1000 + PERAKENDE_SABIT.FIRE_M;

  const frameCost =
    n.framePriceTL > 0 && alanM2 > 0 ? metreToplam * n.framePriceTL : 0;

  const paspartu1 = hasMat ? alanM2 * n.matPrice : 0;
  const paspartu2 = n.innerMatPrice > 0 ? alanM2 * n.innerMatPrice : 0;
  const zemin = n.zeminPrice > 0 ? alanM2 * n.zeminPrice : 0;
  const pencereKesim =
    n.pencereSayisi > 1 && hasMat
      ? (n.pencereSayisi - 1) *
        ((PENCERE_UCRETI[n.matPrice] || 0) +
          (n.innerMatPrice > 0 ? PENCERE_UCRETI[n.innerMatPrice] || 0 : 0))
      : 0;
  const matCost = paspartu1 + paspartu2 + zemin + pencereKesim;

  const glassCost = !n.kasa && n.camPrice > 0 ? alanM2 * n.camPrice : 0;

  // Baskı eserin kendi alanından; pencereli düzende her fotoğraf ayrı baskı
  const printCost =
    n.printUsdPerM2 > 0
      ? (n.wMM / 1000) * (n.hMM / 1000) * n.pencereSayisi * n.printUsdPerM2 * n.usdRate
      : 0;

  return {
    icEn,
    icBoy,
    alanM2,
    metreToplam,
    frameCost,
    matCost,
    pencereKesim,
    glassCost,
    printCost,
    itemTotal: frameCost + matCost + glassCost + printCost,
    yerlesim: alan.yerlesim,
  };
}

export interface PerakendeHata {
  kod: string;
  mesaj: string;
  seviye: "engel" | "uyari";
}

/**
 * Sipariş kuralları — sihirbaz (uyarı/engel) ve sunucu (ret) aynı listeyi
 * kullanır. "engel": fiziksel olarak üretilemez; "uyari": alınabilir ama
 * personel onayı ister (ör. 2 mm gerçek camda büyük boy kırılma riski).
 */
export function dogrulaPerakende(g: PerakendeGirdi): PerakendeHata[] {
  const S = PERAKENDE_SABIT;
  const n = normalize(g);
  const h: PerakendeHata[] = [];
  const push = (kod: string, mesaj: string, seviye: "engel" | "uyari" = "engel") =>
    h.push({ kod, mesaj, seviye });

  if (!(n.wMM > 0 && n.hMM > 0)) {
    push("OLCU_YOK", "Eserin genişliğini ve yüksekliğini girin.");
    return h;
  }
  if (n.wMM < S.MIN_ART_MM || n.hMM < S.MIN_ART_MM)
    push("MIN_OLCU", `Eser ölçüsü en az ${S.MIN_ART_MM} mm olmalı.`);
  if (n.wMM > S.MAX_OUTER_MM || n.hMM > S.MAX_OUTER_MM)
    push("MAX_OLCU", `Eser ölçüsü ${S.MAX_OUTER_MM / 10} cm'i aşamaz.`);

  if (n.pencereSayisi > S.MAX_WINDOWS)
    push("PENCERE", `Pencere sayısı 1–${S.MAX_WINDOWS} arasında olmalı.`);
  if (n.icSeritMm < 0 || n.icSeritMm > S.MAX_MOUNTING_MM)
    push("IC_SERIT", `İç paspartu şeridi 0–${S.MAX_MOUNTING_MM} mm arasında olmalı.`);

  if (n.matPrice > 0) {
    const kenarlar = [n.kenar.ust, n.kenar.alt, n.kenar.sol, n.kenar.sag];
    if (kenarlar.some((v) => v < S.MIN_MAT_EDGE_MM))
      push("PASPARTU_KENAR", `Paspartu kenarı en az ${S.MIN_MAT_EDGE_MM} mm olmalı.`);
    if (kenarlar.some((v) => v > S.MAX_MAT_EDGE_MM))
      push("PASPARTU_KENAR", `Paspartu kenarı en fazla ${S.MAX_MAT_EDGE_MM} mm olabilir.`);
  }

  const sonuc = hesaplaPerakende(g);
  const kisa = Math.min(sonuc.icEn, sonuc.icBoy);
  const uzun = Math.max(sonuc.icEn, sonuc.icBoy);

  if (n.matPrice > 0 && (kisa > S.MAT_SHEET_SHORT_MM || uzun > S.MAT_SHEET_LONG_MM))
    push(
      "PASPARTU_TABAKA",
      "Paspartu dış ölçüsü 80×120 cm tabakayı aşıyor — bu ölçüde paspartu kesilemez."
    );

  if (n.camPrice > 0) {
    const maxKisa = num(g.camMaxKisaMM, S.GLASS_MAX_SHORT_MM) || S.GLASS_MAX_SHORT_MM;
    const maxUzun = num(g.camMaxUzunMM, S.GLASS_MAX_LONG_MM) || S.GLASS_MAX_LONG_MM;
    if (kisa > maxKisa || uzun > maxUzun) {
      push(
        "CAM_OLCU",
        `Bu ölçüde bu cam üretilemez (en büyük ${maxKisa / 10}×${maxUzun / 10} cm) — PVC cam seçin.`
      );
    } else if (
      num(g.camUyariKisaMM) > 0 &&
      (kisa > num(g.camUyariKisaMM) || uzun > num(g.camUyariUzunMM, num(g.camUyariKisaMM)))
    ) {
      push(
        "CAM_RISK",
        `Gerçek cam ${num(g.camUyariKisaMM) / 10}×${num(g.camUyariUzunMM) / 10} cm üstünde kırılganda — kargolanacaksa kırılmayan (PVC) cam önerin.`,
        "uyari"
      );
    }
  }

  if (sonuc.icEn > S.MAX_OUTER_MM || sonuc.icBoy > S.MAX_OUTER_MM)
    push(
      "DIS_OLCU",
      `Çerçevenin ölçüsü ${S.MAX_OUTER_MM / 10} cm'i aşıyor — tek parça üretilemez.`
    );

  return h;
}
