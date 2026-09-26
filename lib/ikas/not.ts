// Web sitesindeki çerçeve hesaplayıcının ikas siparişine yazdığı
// "CERCEVE SIPARIS DETAYI (OZL-…)" metnini üretim kalemine çevirir.
// KURAL TABANLI; yapay zekâ gerekmez. Metin sipariş notunda, sipariş
// özniteliklerinde ya da satır seçeneklerinde olabilir — hepsi taranır.
// Not yoksa varyant adından ("Özel Çerçeve: 1266-02 – Eser 50×70 cm (Dış ≈54×74 cm) – Kırılmayan Mat Cam")
// SKU + eser ölçüsü + cam okunur.
//
// Örnek blok (tek satır ya da satırlara bölünmüş gelebilir):
//   CERCEVE SIPARIS DETAYI (OZL-mugu66is-6053) Urun: Özel Çerçeve: 1266-02 – Eser 50×70 cm … | Adet: 3 [URUN]
//   • SKU (cerceve profili): 1266-02 | Adet: 3 | Profil genisligi: 20 mm | Icerik: Diploma / Belge • Yon: Dikey [OLCULER (mm)]
//   • Sanat eseri: 500 × 700 mm … • Kesim: 502 × 702 mm … • Cerceve ic olcusu (paspartu dahil): 500 × 700 mm …
//   • Tahmini dis olcu (cerceve profili dahil): ~540 × 740 mm + 2mm pay: Eklendi [PASPARTU] • Paspartu: Yok [CAM]
//   • Cam: Kırılmayan Mat Cam | Cam olcusu: … [FIYAT] • Fiyat: Cerceve 2.097,90 TL • Paspartu 0,00 TL • Cam 2.100,00 TL • TOPLAM: 4.197,90 TL
//   • SUNUCU DOGRULAMASI: metre fiyati 259,00 TL/m (sayfa), tutar bu fiyatla yeniden hesaplandi

import type { RetailItem } from "@/lib/retail-orders";
import { MAT_TYPES } from "@/data/perakende";
import type { IsKalem } from "@/lib/uretim/tur";
import { olcuCm } from "@/lib/uretim/tur";
import type { IkasOrder, IkasSatir } from "./client";

// Türkçe harfleri UZUNLUĞU KORUYARAK ASCII'ye indirger: eşleşme katlanmış metinde
// yapılır, değer asıl metinden aynı konumlardan alınır.
const KATLA: Record<string, string> = { ç: "c", Ç: "c", ğ: "g", Ğ: "g", ı: "i", I: "i", İ: "i", ö: "o", Ö: "o", ş: "s", Ş: "s", ü: "u", Ü: "u" };
export function katla(s: string): string {
  let out = "";
  for (const ch of s) out += KATLA[ch] ?? ch.toLowerCase();
  return out;
}

/** "1.916,60" → 1916.6 · "1000" → 1000 · "104,5" → 104.5 */
export function sayi(s: string | undefined | null): number {
  if (!s) return 0;
  let t = String(s).trim().replace(/\s/g, "");
  if (/,\d{1,2}$/.test(t)) t = t.replace(/\./g, "").replace(",", ".");
  else if (/^\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, "");
  else t = t.replace(",", ".");
  const n = Number(t);
  return Number.isFinite(n) ? n : 0;
}

export interface CozulenKalem {
  ozl: string;
  sku: string;
  adet: number;
  profilMm?: number;
  icerik?: string;
  yon?: string;
  eserMm?: { w: number; h: number };
  kesimMm?: { w: number; h: number };
  icMm?: { w: number; h: number };     // çerçeve iç ölçüsü (paspartu dahil)
  disMm?: { w: number; h: number };
  pay?: boolean;
  paspartu?: string;                    // "Yok" ya da metin
  paspartuKenarMm?: { ust: number; sag: number; alt: number; sol: number };
  cam?: string;
  camOlcu?: string;
  fiyat?: { cerceve: number; paspartu: number; cam: number; toplam: number };
  metreFiyat?: number;
  kaynak: "not" | "varyant";
}

const OLCU = "([\\d.,]+)\\s*[x×*]\\s*([\\d.,]+)";

function olcuBul(k: string, etiket: RegExp): { w: number; h: number } | undefined {
  const m = etiket.exec(k);
  if (!m) return undefined;
  const w = sayi(m[1]), h = sayi(m[2]);
  return w > 0 && h > 0 ? { w, h } : undefined;
}

/** Katlanmış metinde etiket sonrası değeri (sonraki ayraca kadar) asıl metinden alır. */
function deger(asil: string, k: string, etiket: RegExp): string | undefined {
  const m = etiket.exec(k);
  if (!m || m.index === undefined) return undefined;
  const bas = m.index + m[0].length;
  const kalan = k.slice(bas);
  const son = kalan.search(/[|•\[\n]|\s{2,}|(?=\s(?:adet|profil genisligi|icerik|yon|sku)\s*[:(])/);
  const parca = asil.slice(bas, son >= 0 ? bas + son : undefined).trim();
  return parca.replace(/[\s:.-]+$/, "").trim() || undefined;
}

/** Metni "CERCEVE SIPARIS DETAYI" bloklarına ayırır (ozl kimliğiyle). */
export function notBloklari(metin: string): { ozl: string; metin: string }[] {
  if (!metin) return [];
  const k = katla(metin);
  const re = /cerceve siparis detayi\s*\(?\s*(ozl-[a-z0-9-]+)?/g;
  const baslar: { i: number; ozl: string }[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(k))) baslar.push({ i: m.index, ozl: m[1] ? metin.slice(m.index + m[0].indexOf(m[1]), m.index + m[0].indexOf(m[1]) + m[1].length) : "" });
  if (!baslar.length) return [];
  return baslar.map((b, idx) => ({ ozl: b.ozl.toUpperCase(), metin: metin.slice(b.i, idx + 1 < baslar.length ? baslar[idx + 1].i : undefined) }));
}

/** Tek bloğu çözer; SKU ya da eser ölçüsü okunamazsa null. */
export function notCoz(blok: string): CozulenKalem | null {
  const asil = blok;
  const k = katla(blok);
  const ozl = /ozl-[a-z0-9-]+/.exec(k);
  const sku = deger(asil, k, /sku\s*(?:\([^)]*\))?\s*:\s*/);
  const adetM = /adet\s*:\s*(\d+)/.exec(k);
  const eser = olcuBul(k, new RegExp(`sanat eseri\\s*:\\s*${OLCU}\\s*mm`));
  const kesim = olcuBul(k, new RegExp(`kesim\\s*:\\s*${OLCU}\\s*mm`));
  const ic = olcuBul(k, new RegExp(`ic olcusu[^:]*:\\s*${OLCU}\\s*mm`));
  const dis = olcuBul(k, new RegExp(`dis olcu[^:]*:\\s*~?\\s*${OLCU}\\s*mm`));
  if (!sku && !eser) return null;
  const profilM = /profil genisligi\s*:\s*([\d.,]+)\s*mm/.exec(k);
  const payM = /2\s*mm\s*pay\s*:\s*(eklendi|eklenmedi)/.exec(k);
  const yonM = /\byon\s*:\s*(yatay|dikey|kare)/.exec(k);
  const icerik = deger(asil, k, /icerik\s*:\s*/);
  const paspartuHam = deger(asil, k, /paspartu\s*:\s*(?!\s*[\d.,]+\s*tl)/);
  const camHam = deger(asil, k, /\bcam\s*:\s*(?!\s*[\d.,]+\s*tl)/);
  const camOlcu = deger(asil, k, /cam olcusu\s*:\s*/);
  const fiyatM = /cerceve\s+([\d.,]+)\s*tl\s*[•|,;]?\s*paspartu\s+([\d.,]+)\s*tl\s*[•|,;]?\s*cam\s+([\d.,]+)\s*tl\s*[•|,;]?\s*toplam\s*:?\s*([\d.,]+)\s*tl/.exec(k);
  const toplamM = /toplam\s*:\s*([\d.,]+)\s*tl/.exec(k);
  const metreM = /metre fiyati\s+([\d.,]+)\s*tl/.exec(k);

  const out: CozulenKalem = {
    ozl: ozl ? asil.slice(ozl.index, ozl.index + ozl[0].length).toUpperCase() : "",
    sku: (sku || "").replace(/\s+/g, " ").trim(),
    adet: adetM ? Math.max(1, Number(adetM[1])) : 1,
    kaynak: "not",
  };
  if (profilM) out.profilMm = sayi(profilM[1]);
  if (icerik && !/^belirtilmedi$/i.test(katla(icerik))) out.icerik = icerik;
  if (yonM) out.yon = yonM[1] === "yatay" ? "Yatay" : yonM[1] === "dikey" ? "Dikey" : "Kare";
  if (eser) out.eserMm = eser;
  if (kesim) out.kesimMm = kesim;
  if (ic) out.icMm = ic;
  if (dis) out.disMm = dis;
  if (payM) out.pay = payM[1] === "eklendi";
  if (paspartuHam) {
    const pk = katla(paspartuHam);
    if (/^yok\b/.test(pk)) out.paspartu = "Yok";
    else {
      out.paspartu = paspartuHam;
      const mmler = [...pk.matchAll(/([\d.,]+)\s*mm/g)].map((x) => sayi(x[1])).filter((n) => n > 0);
      const cmler = mmler.length ? [] : [...pk.matchAll(/([\d.,]+)\s*cm/g)].map((x) => sayi(x[1]) * 10).filter((n) => n > 0);
      const nums = mmler.length ? mmler : cmler;
      if (nums.length >= 4) out.paspartuKenarMm = { ust: nums[0], sag: nums[1], alt: nums[2], sol: nums[3] };
      else if (nums.length === 2) out.paspartuKenarMm = { ust: nums[0], alt: nums[0], sag: nums[1], sol: nums[1] };
      else if (nums.length === 1) out.paspartuKenarMm = { ust: nums[0], sag: nums[0], alt: nums[0], sol: nums[0] };
      else if (ic && eser && (ic.w > eser.w || ic.h > eser.h)) {
        const y = Math.max(0, Math.round((ic.w - eser.w) / 2)), d = Math.max(0, Math.round((ic.h - eser.h) / 2));
        out.paspartuKenarMm = { ust: d, alt: d, sag: y, sol: y };
      }
    }
  }
  if (camHam) out.cam = /^yok\b/.test(katla(camHam)) ? "Yok" : camHam;
  if (camOlcu) out.camOlcu = camOlcu;
  if (fiyatM) out.fiyat = { cerceve: sayi(fiyatM[1]), paspartu: sayi(fiyatM[2]), cam: sayi(fiyatM[3]), toplam: sayi(fiyatM[4]) };
  else if (toplamM) out.fiyat = { cerceve: 0, paspartu: 0, cam: 0, toplam: sayi(toplamM[1]) };
  if (metreM) out.metreFiyat = sayi(metreM[1]);
  return out;
}

/** Varyant adından: "Özel Çerçeve: 1266-02 – Eser 50×70 cm (Dış ≈54×74 cm) – Kırılmayan Mat Cam" */
export function varyantAdiCoz(ad: string, sku?: string | null): CozulenKalem | null {
  const asil = String(ad || "");
  const k = katla(asil);
  // SKU: "Çerçeve:" sonrası, " – " (boşluklu tire) ya da satır sonuna kadar; koddaki tire ("1266-02", "KS3420-BLACK") korunur
  const skuM = /cerceve\s*:\s*([a-z0-9][a-z0-9.\/-]*(?: [a-z0-9][a-z0-9.\/-]*)*?)(?=\s+[–—-]\s+|\s*$|\s*\()/.exec(k);
  const eserM = new RegExp(`eser\\s*${OLCU}\\s*cm`).exec(k);
  const disM = new RegExp(`dis\\s*[≈~]?\\s*${OLCU}\\s*cm`).exec(k);
  const parcalar = asil.split(/\s[–—-]\s/);
  const cam = parcalar.length >= 3 ? parcalar[parcalar.length - 1].trim() : undefined;
  const skuTxt = skuM ? asil.slice(skuM.index + skuM[0].indexOf(skuM[1]), skuM.index + skuM[0].indexOf(skuM[1]) + skuM[1].length).trim() : "";
  if (!skuTxt && !eserM) return null;
  const out: CozulenKalem = { ozl: (sku || "").toUpperCase(), sku: skuTxt, adet: 1, kaynak: "varyant" };
  if (eserM) { const w = sayi(eserM[1]) * 10, h = sayi(eserM[2]) * 10; if (w > 0 && h > 0) out.eserMm = { w, h }; }
  if (disM) { const w = sayi(disM[1]) * 10, h = sayi(disM[2]) * 10; if (w > 0 && h > 0) out.disMm = { w, h }; }
  if (cam && /cam|pleksi|pvc/i.test(katla(cam))) out.cam = cam;
  return out;
}

function matAdi(paspartu?: string): { matType: string; matCode: string } {
  if (!paspartu || /^yok\b/.test(katla(paspartu))) return { matType: "Paspartu Yok", matCode: "-" };
  const pk = katla(paspartu);
  const bulunan = MAT_TYPES.find((m) => m.code !== "-" && pk.includes(katla(m.name)));
  return bulunan ? { matType: bulunan.name, matCode: bulunan.code } : { matType: paspartu.slice(0, 60), matCode: "-" };
}

const r2 = (n: number) => Math.round((Number(n) || 0) * 100) / 100;

/** Çözülen kalemi üretim kalemine (föy verisiyle) çevirir. */
export function kalemeCevir(c: CozulenKalem, satir?: { adet?: number; fiyat?: number }): IsKalem {
  const adet = Math.max(1, c.adet || satir?.adet || 1);
  const eser = c.eserMm || { w: 0, h: 0 };
  const kenar = c.paspartuKenarMm || { ust: 0, sag: 0, alt: 0, sol: 0 };
  const { matType, matCode } = matAdi(c.paspartu);
  const cam = c.cam && !/^yok\b/.test(katla(c.cam)) ? c.cam : "Cam Yok";
  const toplam = c.fiyat?.toplam || satir?.fiyat || 0;
  const bol = (v: number) => r2(v / adet);
  const retail: RetailItem | undefined = eser.w > 0 && eser.h > 0 ? {
    artWidth: eser.w, artWidthUnit: "mm", artHeight: eser.h, artHeightUnit: "mm",
    frameCode: c.sku || "OZEL",
    framePriceTL: c.metreFiyat || 0,
    manualPrice: false,
    matType, matCode, matColor: "-", matColorHex: "-",
    doubleMat: false, innerMatType: "-", innerMatColor: "-", innerMatColorHex: "-", altMontaj: "-",
    zeminEnabled: false, zeminType: "-", zeminColor: "-", zeminColorHex: "-",
    matTop: kenar.ust, matRight: kenar.sag, matBottom: kenar.alt, matLeft: kenar.sol,
    pencereSayisi: 1, pencereDuzen: "", pencereAralik: 0, kasa: false,
    glassType: cam.slice(0, 60), printType: "Baskı Yok",
    frameCost: bol(c.fiyat?.cerceve || 0), matCost: bol(c.fiyat?.paspartu || 0), glassCost: bol(c.fiyat?.cam || 0), printCost: 0,
    itemTotal: bol(toplam),
  } : undefined;
  const parcalar: string[] = [];
  if (eser.w > 0) parcalar.push(`Eser ${olcuCm(eser)}`);
  if (c.disMm) parcalar.push(`Dış ~${olcuCm(c.disMm)}`);
  parcalar.push(matType === "Paspartu Yok" ? "Paspartu yok" : `Paspartu: ${c.paspartu}`);
  parcalar.push(cam === "Cam Yok" ? "Cam yok" : cam);
  if (c.pay !== undefined) parcalar.push(c.pay ? "+2 mm pay" : "pay yok");
  if (c.icerik) parcalar.push(c.icerik);
  const out: IsKalem = { sku: c.sku || "OZEL", adet, ozet: parcalar.join(" · ") };
  if (eser.w > 0) out.eserMm = eser;
  if (c.disMm) out.disMm = c.disMm;
  if (c.icerik) out.icerik = c.icerik;
  if (c.yon) out.yon = c.yon;
  out.paspartu = matType === "Paspartu Yok" ? "Yok" : c.paspartu;
  out.cam = cam;
  if (c.pay !== undefined) out.pay = c.pay;
  if (toplam) out.fiyat = r2(toplam);
  if (retail) out.retail = retail;
  return out;
}

/** Siparişteki not/öznitelik/seçenek metinlerini tek metinde toplar. */
export function siparisMetinleri(o: IkasOrder): string {
  const parcalar: string[] = [];
  if (o.note) parcalar.push(String(o.note));
  for (const a of o.attributes || []) if (a?.value) parcalar.push(String(a.value));
  for (const s of o.orderLineItems || []) for (const op of s.options || []) for (const v of op?.values || []) if (v?.value) parcalar.push(String(v.value));
  return parcalar.join("\n");
}

export interface KalemCozum { kalemler: IsKalem[]; notMetni: string; eksik: string[] }

/**
 * Siparişin satırlarını üretim kalemlerine çevirir. Her satır için önce SKU'suyla
 * (OZL-…) eşleşen not bloğu, sonra sırayla kalan bloklar, en son varyant adı denenir.
 * `ekMetin`: zaman çizelgesi gibi API'de olmayan kaynaklardan gelen metin (isteğe bağlı).
 */
export function ikasSiparisKalemleri(o: IkasOrder, ekMetin = ""): KalemCozum {
  const notMetni = [siparisMetinleri(o), ekMetin].filter(Boolean).join("\n");
  const bloklar = notBloklari(notMetni).map((b) => ({ ...b, coz: notCoz(b.metin), kullanildi: false }));
  const kalemler: IsKalem[] = [];
  const eksik: string[] = [];
  for (const satir of o.orderLineItems || []) {
    const adet = Math.max(1, Math.round(Number(satir.quantity) || 1));
    const satirFiyat = Number(satir.finalPrice ?? satir.price) || 0;
    const skuUp = String(satir.variant?.sku || "").toUpperCase();
    let coz: CozulenKalem | null = null;
    let b = skuUp ? bloklar.find((x) => !x.kullanildi && x.coz && x.ozl && x.ozl === skuUp) : undefined;
    if (!b) b = bloklar.find((x) => !x.kullanildi && x.coz && !x.ozl);
    if (!b && bloklar.length === (o.orderLineItems || []).length) b = bloklar.find((x) => !x.kullanildi && x.coz);
    if (b) { b.kullanildi = true; coz = b.coz; }
    if (!coz) coz = varyantAdiCoz(satir.variant?.name || "", satir.variant?.sku);
    if (!coz) {
      eksik.push(satir.variant?.name || satir.id);
      kalemler.push({ sku: satir.variant?.sku || "?", adet, ozet: (satir.variant?.name || "Kalem").slice(0, 300), fiyat: r2(satirFiyat) });
      continue;
    }
    if (coz.kaynak === "not" && coz.adet !== adet && !coz.fiyat) coz.adet = adet;
    if (coz.kaynak === "varyant") coz.adet = adet;
    kalemler.push(kalemeCevir(coz, { adet, fiyat: satirFiyat }));
  }
  return { kalemler, notMetni, eksik };
}

export type { IkasSatir };
