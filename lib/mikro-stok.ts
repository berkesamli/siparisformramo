// Mikro'dan stok çekme — günlük Excel yüklemesinin karşılığı, AYNI formüllerle:
//   • yalnızca adında PROFİL geçen kalemler,
//   • kod = stok adının tanım kelimesine (LAMİNE / PROFİL …) kadar olan kısmı
//     (extractCode — Excel ayrıştırıcısıyla birebir aynı fonksiyon),
//   • depo adında ANKARA / İSTANBUL geçenler ilgili şubeye toplanır,
//   • miktar metre olarak yayınlanır; ekran 2,9'a bölerek "boy" gösterir.
// Sonuç Excel yüklemesiyle aynı StockData biçimindedir; aynı Blob dosyasına yazılır.

import { stokMiktarlari, mikroConfigured, type MikroStokSonuc } from "./mikro";
import { extractCode, norm, type StockData, type StockItem, stokAnahtar } from "./stock-parse";
import { STOCK_BLOB_PATH, stockBlobConfigured } from "./stock-store";

export interface MikroStokBilgi {
  depolar: { no: number; ad: string }[];
  ankaraDepolar: number[];
  istanbulDepolar: number[];
  yontem?: string;
  birimler: string[];
  hamSatir: number;   // Mikro'dan gelen satır
  profilSatir: number; // adında PROFİL geçen satır
  kalem: number;      // kod bazında kalem
  sureMs: number;
  ornek: { kod: string; isim: string; birim: string; ankara: number; istanbul: number; cikarilanKod: string }[];
}

export interface MikroStokCekim {
  ok: boolean;
  hata?: string;
  data?: StockData;
  bilgi: MikroStokBilgi;
}

function bilgiOlustur(r: MikroStokSonuc, profilSatir: number, kalem: number): MikroStokBilgi {
  return {
    depolar: r.depolar,
    ankaraDepolar: r.ankaraDepolar,
    istanbulDepolar: r.istanbulDepolar,
    yontem: r.yontem,
    birimler: Array.from(new Set(r.satirlar.map((s) => s.birim).filter(Boolean))).slice(0, 6),
    hamSatir: r.satirlar.length,
    profilSatir,
    kalem,
    sureMs: r.sureMs,
    ornek: r.satirlar
      .filter((s) => norm(s.isim).includes("PROFIL") && (s.ankara > 0 || s.istanbul > 0))
      .slice(0, 8)
      .map((s) => ({ ...s, cikarilanKod: extractCode(s.isim) })),
  };
}

export async function mikroStokCek(): Promise<MikroStokCekim> {
  const r = await stokMiktarlari();
  if (!r.ok) return { ok: false, hata: r.hata, bilgi: bilgiOlustur(r, 0, 0) };

  const map = new Map<string, StockItem>();
  let profilSatir = 0;
  for (const s of r.satirlar) {
    if (!norm(s.isim).includes("PROFIL")) continue; // Excel'deki süzgeç
    profilSatir++;
    const code = extractCode(s.isim);
    if (!code) continue;
    // Excel'de miktarı 0 veya negatif olan depo satırları atlanıyordu
    const ank = s.ankara > 0 ? s.ankara : 0;
    const ist = s.istanbul > 0 ? s.istanbul : 0;
    if (ank <= 0 && ist <= 0) continue;
    // Aynı ürünün iki yazımı (KS4022-BIG / KS4022-BİG) tek kalemde toplanır
    const anahtar = stokAnahtar(code);
    let item = map.get(anahtar);
    if (!item) { item = { code, ankaraMt: 0, istanbulMt: 0 }; map.set(anahtar, item); }
    item.ankaraMt += ank;
    item.istanbulMt += ist;
  }
  const items = Array.from(map.values())
    .map((it) => ({ code: it.code, ankaraMt: Math.round(it.ankaraMt * 100) / 100, istanbulMt: Math.round(it.istanbulMt * 100) / 100 }))
    .sort((a, b) => a.code.localeCompare(b.code, "tr"));
  const bilgi = bilgiOlustur(r, profilSatir, items.length);
  if (!items.length) {
    return { ok: false, hata: `Mikro'dan ${r.satirlar.length} satır geldi ama adında PROFİL geçen ve miktarı olan kalem yok.`, bilgi };
  }
  const simdi = new Date();
  const data: StockData = {
    updatedAt: simdi.toISOString(),
    sourceName: `Mikro (canlı) · ${simdi.toLocaleString("tr-TR", { timeZone: "Europe/Istanbul", day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" })}`,
    items,
  };
  return { ok: true, data, bilgi };
}

/** Mikro'dan çek ve Blob'a yaz (Excel yüklemesiyle aynı dosya). */
export async function mikroStokCekVeKaydet(): Promise<MikroStokCekim & { kaydedildi: boolean }> {
  if (!mikroConfigured()) {
    return { ok: false, hata: "Mikro bağlantısı ayarlanmamış.", kaydedildi: false, bilgi: { depolar: [], ankaraDepolar: [], istanbulDepolar: [], birimler: [], hamSatir: 0, profilSatir: 0, kalem: 0, sureMs: 0, ornek: [] } };
  }
  const r = await mikroStokCek();
  if (!r.ok || !r.data) return { ...r, kaydedildi: false };
  if (!stockBlobConfigured()) return { ...r, ok: false, hata: "Kalıcı depolama (Blob) bağlı değil; stok kaydedilemedi.", kaydedildi: false };
  const { put } = await import("@vercel/blob");
  await put(STOCK_BLOB_PATH, JSON.stringify(r.data), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return { ...r, kaydedildi: true };
}

/** Yayındaki stok bu kadar dakikadan eskiyse Mikro'dan tazelenir (STOK_TAZELIK_DK, varsayılan 120). */
export function stokTazelikDk(): number {
  const n = parseInt(process.env.STOK_TAZELIK_DK || "", 10);
  return Number.isInteger(n) && n > 0 ? n : 120;
}
