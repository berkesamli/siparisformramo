// Çerçeve metre fiyatı: toptan liste (USD/mt) × bayi çarpanı × kur.
import { findProfile } from "@/data/catalog";

export interface FramePriceResult {
  found: boolean;
  code: string;
  resolvedCode?: string;
  series?: string;
  tlPerM: number; // bayinin satış fiyatı
  costTlPerM: number; // toptan liste maliyeti (KDV hariç) — bayiye bilgi
}

export function dealerFramePrice(code: string, usdRate: number, factor: number): FramePriceResult {
  const profile = findProfile(code);
  if (!profile || !(usdRate > 0) || !(factor > 0)) {
    return { found: false, code, tlPerM: 0, costTlPerM: 0 };
  }
  const r2 = (n: number) => Math.round(n * 100) / 100;
  return {
    found: true,
    code,
    resolvedCode: profile.code,
    series: profile.series,
    tlPerM: r2(profile.priceUSD * factor * usdRate),
    costTlPerM: r2(profile.priceUSD * usdRate),
  };
}
