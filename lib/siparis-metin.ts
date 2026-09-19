// Sipariş metni çözümleyici — KURAL TABANLI (yapay zekâ gerekmez).
//
// "Metinden Sipariş Oluştur" kutusuna yapıştırılan satırları forma dökülecek
// satırlara çevirir. Yapay zekâ yalnızca burada okunamayan satırlar için
// (route içinde) devreye girer. Kurallar:
//   • "KS 3420-black 10 koli", "GB211-4110B 3 koli", "3127 S-A79 20 boy",
//     "NS 455 → 25 ad", "3 koli ks2030 beyaz", "gc065 1473 50 mt" okunur.
//   • Çerçeve kodu = katalogdaki ana kod (boşluksuz, stok listesindeki gibi:
//     "KS3420", "3127S") + "-" + desen/renk eki BÜYÜK HARF ("KS3420-BLACK").
//     Ek hiçbir zaman atılmaz.
//   • El yazısından gelen "6B 211" gibi kodlar katalogda yoksa "GB211" denenir.
//   • Teknik malzeme ürün adıyla bulunur; "NS 455" → NS Karton Düz, karton kodu 455.
//   • "%40 isk + KDV", "Araç ile gidecek" gibi başlık satırları iskonto / KDV /
//     not alanlarına gider; ilk ürün olmayan satır müşteri adıdır.

import { profilBul, type FrameProfile, type TechnicalProduct } from "./catalog-utils";

export type MetinTur = "frame" | "glass" | "ayna" | "technical" | "other";

export interface MetinSatir {
  kind: MetinTur;
  code: string;        // forma yazılacak kod / ürün adı
  rawCode: string;     // metinde yazan hali
  matched: boolean;    // katalogda bulundu
  unit: string;        // koli | boy | metre | adet | kutu | plaka
  qty: number;
  note: string;
  confidence: number;  // 1 birebir, 0.7 büyük olasılıkla, 0.4 kontrol edilmeli
  techCode?: string;   // teknik ürün kodu (technical.ts)
  kartonKodu?: string; // karton ürünlerinde renk/desen kodu
}

export interface MetinSonuc {
  customer: string;
  note: string;
  iskontoPct?: number;
  kdv?: boolean;
  lines: MetinSatir[];
  kalan: string[];     // kural tabanlı okunamayan satırlar (yapay zekâya gider)
}

const BIRIM: Record<string, string> = {
  koli: "koli", kolı: "koli", boy: "boy", metre: "metre", mt: "metre", m: "metre",
  adet: "adet", ad: "adet", kutu: "kutu", paket: "kutu", plaka: "plaka",
};
const BIRIM_RE = "(koli|kolı|boy|metre|mt|m|adet|ad|kutu|paket|plaka)";
const MIKTAR_RE = "(\\d+(?:[.,]\\d+)?)";
// "<kod> [→ : - x] <miktar> <birim>"
const SONDA = new RegExp(`^(.*?)\\s*(?:[→>:=]|-{1,2}>|x|×)?\\s*${MIKTAR_RE}\\s*${BIRIM_RE}\\.?\\s*$`, "i");
// "<miktar> <birim> <kod>"
const BASTA = new RegExp(`^${MIKTAR_RE}\\s*${BIRIM_RE}\\.?\\s+(.+?)\\s*$`, "i");
// "<kod> <miktar>" — birim yazılmamış (çerçevede koli varsayılır, güven düşük;
// 500'den büyük sayı miktar değil kod parçası sayılır)
const BIRIMSIZ = new RegExp(`^(.*?[A-Za-z0-9)])(?:\\s*(?:[→>:=]|x|×)\\s*|\\s+)(\\d{1,3})\\s*$`);

const norm = (s: string) =>
  String(s || "")
    .toLocaleUpperCase("tr-TR")
    .replace(/İ/g, "I")
    .replace(/['’`´.]/g, "")
    .replace(/[^A-Z0-9ÇĞÖŞÜ ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const sayi = (s: string) => Number(String(s).replace(",", ".")) || 0;

/** Katalog kodunu stok listesindeki yazıma çevirir: "KS 3420" → "KS3420", "3127 S" → "3127S". */
export const stokBicimi = (catalogCode: string) => catalogCode.replace(/\s+/g, "");

/**
 * Çerçeve kodunu çözer: ana kod katalogdan, desen/renk eki metinden.
 * "KS 3420 - black" → KS3420-BLACK · "6B 211-4110B" → GB211-4110B ·
 * "gc065 1473" → GC065-1473 · "3127 S-A79" → 3127S-A79
 */
export function cerceveKodu(raw: string, profiles: FrameProfile[]): { code: string; matched: boolean; note?: string } {
  const s = raw.trim().replace(/[–—]/g, "-").replace(/\s*-\s*/g, "-").replace(/\s+/g, " ").replace(/^-+|-+$/g, "");
  const parcalar = s.split("-");
  let base = parcalar[0];
  let ek = parcalar.slice(1).join("-");
  let p = profilBul(profiles, base);
  let not: string | undefined;
  if (!p && !ek) {
    // Tire yerine boşluk: "gc065 1473", "ks3420 black"
    const m = s.match(/^(\S+)\s+(.+)$/);
    if (m) {
      const pp = profilBul(profiles, m[1]);
      if (pp) { p = pp; base = m[1]; ek = m[2]; }
    }
  }
  if (!p && /^6[A-Za-z]/.test(base)) {
    // El yazısında G harfi 6 gibi okunur: 6B 211 → GB211
    const pp = profilBul(profiles, "G" + base.slice(1));
    if (pp) { p = pp; not = `"${base}" katalogda yok, ${stokBicimi(pp.code)} olarak alındı`; }
  }
  const ekTemiz = ek.trim().replace(/\s+/g, " ").toUpperCase().replace(/İ/g, "I");
  if (!p) return { code: s.toUpperCase().replace(/İ/g, "I"), matched: false };
  return { code: stokBicimi(p.code) + (ekTemiz ? "-" + ekTemiz : ""), matched: true, note: not };
}

/** Teknik ürünü adından bulur; karton ürünlerinde kalan kod parçasını karton kodu olarak döndürür. */
export function teknikUrun(raw: string, technical: TechnicalProduct[]): { t: TechnicalProduct; kartonKodu?: string } | null {
  const n = norm(raw);
  if (!n) return null;
  // "NS 455", "NS455", "NS 413A" → NS Karton Düz + karton kodu
  const ns = n.match(/^NS\s*([0-9]{2,4}[A-Z]?)$/);
  if (ns) {
    const t = technical.find((x) => x.code === "NS-KARTON-DUZ") || technical.find((x) => norm(x.name).startsWith("NS KARTON"));
    if (t) return { t, kartonKodu: ns[1] };
  }
  if (/OLUKLU/.test(n)) {
    const t = technical.find((x) => /OLUKLU/.test(norm(x.name)));
    if (t) return { t };
  }
  const satirTok = new Set(n.split(" ").filter(Boolean));
  let enIyi: { t: TechnicalProduct; puan: number; kalan: string[] } | null = null;
  for (const t of technical) {
    const adTok = norm(t.name).split(" ").filter((x) => x.length >= 2);
    if (!adTok.length) continue;
    const ortak = adTok.filter((x) => satirTok.has(x)).length;
    if (ortak === 0) continue;
    // Ürün adının en az yarısı ve en az 2 kelimesi (tek kelimeli adlarda 1) satırda geçmeli
    const gerekli = Math.max(1, Math.min(2, adTok.length), Math.ceil(adTok.length / 2));
    if (ortak < gerekli) continue;
    const eksik = adTok.length - ortak;
    let puan = ortak - 0.15 * eksik;
    if (t.category === "OLGA") puan += 0.05; // belirsizlikte kendi ürünümüz
    // Marka adı satırda geçiyorsa o ürün (alfa / cassese / ithal / kartuşlu)
    const kategori = norm(t.category);
    if (kategori && satirTok.has(kategori.split(" ")[0])) puan += 0.5;
    if (!enIyi || puan > enIyi.puan) {
      const adSet = new Set(adTok);
      enIyi = { t, puan, kalan: [...satirTok].filter((x) => !adSet.has(x)) };
    }
  }
  if (!enIyi) return null;
  const kartonKodu = enIyi.t.isKarton ? enIyi.kalan.find((x) => /^[0-9]{1,4}[A-Z]{0,2}$/.test(x)) : undefined;
  return { t: enIyi.t, kartonKodu };
}

const BASLIK_ANAHTAR = /ISK|ISKONTO|INDIRIM|KDV|K D V|ARAC|KARGO|TESLIM|GIDECEK|VERILECEK|ACIL|NOT\b|ODEME|PESIN|VADE/;

/** Bir satırın ürün olmayan başlık bilgilerini ayıklar (iskonto, KDV, not). */
function baslikOku(line: string, sonuc: MetinSonuc): boolean {
  const n = norm(line);
  let kullanildi = false;
  const isk = line.match(/%\s*(\d{1,2}(?:[.,]\d)?)|(\d{1,2}(?:[.,]\d)?)\s*%/);
  if (isk && /ISK|INDIRIM|ISKONTO/.test(n)) { sonuc.iskontoPct = sayi(isk[1] || isk[2]); kullanildi = true; }
  else if (/ISK(ONTO)? YOK|INDIRIM YOK|ISKONTOSUZ/.test(n)) { sonuc.iskontoPct = 0; kullanildi = true; }
  if (/KDV|K D V/.test(n)) {
    sonuc.kdv = !/KDV\s*(YOK|SIZ|HARIC DEGIL)|KDVSIZ|FATURASIZ/.test(n);
    kullanildi = true;
  }
  // Not: teslimat / ödeme ifadeleri
  const notParca = line
    .replace(/%\s*\d{1,2}(?:[.,]\d)?\s*(isk(onto)?|indirim)?\.?/gi, "")
    .replace(/\+?\s*k\.?d\.?v\.?/gi, "")
    .replace(/^[-=*#·•\s]+|[-=*#·•\s]+$/g, "")
    .replace(/^\d+[a-z]?\)\s*/i, "")
    .replace(/[·|]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (notParca && kullanildi) {
    sonuc.note = [sonuc.note, notParca].filter(Boolean).join(" · ");
  } else if (notParca && BASLIK_ANAHTAR.test(norm(notParca))) {
    sonuc.note = [sonuc.note, notParca].filter(Boolean).join(" · ");
    kullanildi = true;
  }
  return kullanildi;
}

export function siparisMetniCoz(text: string, profiles: FrameProfile[], technical: TechnicalProduct[]): MetinSonuc {
  const sonuc: MetinSonuc = { customer: "", note: "", lines: [], kalan: [] };
  const satirlar = String(text || "").split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  for (const ham of satirlar) {
    // Ayraç / başlık süsleri
    if (/^[-=*#_~·•]{3,}$/.test(ham)) continue;
    let line = ham.replace(/^[-=*#·•]+\s*|\s*[-=*#·•]+$/g, "").replace(/^\d+[a-z]?\)\s*/i, "").trim();
    if (!line) continue;

    let kod = "", qty = 0, birim = "", guven = 1;
    let m = line.match(SONDA);
    if (m) { kod = m[1]; qty = sayi(m[2]); birim = BIRIM[m[3].toLowerCase()] || "koli"; }
    else if ((m = line.match(BASTA))) { qty = sayi(m[1]); birim = BIRIM[m[2].toLowerCase()] || "koli"; kod = m[3]; }
    else if ((m = line.match(BIRIMSIZ))) { kod = m[1]; qty = sayi(m[2]); birim = ""; guven = 0.6; }

    kod = kod.replace(/[→>:=]+$/g, "").trim();
    if (!kod || qty <= 0) {
      // Ürün satırı değil: başlık / müşteri / not
      if (baslikOku(line, sonuc)) continue;
      const n = norm(line);
      const kelime = line.split(/\s+/).length;
      if (!sonuc.customer && !/\d{3,}/.test(line) && n.length >= 3 && kelime <= 5 && !/^(bir de|birde|ve|ayrıca|lütfen|lutfen|acil|not|merhaba|selam)\b/i.test(line)) {
        sonuc.customer = line.replace(/\s+/g, " ");
        continue;
      }
      sonuc.kalan.push(ham);
      continue;
    }

    const n = norm(kod);
    // 1) Teknik malzeme (agraf, karton, oluklu, NS …)
    const tek = teknikUrun(kod, technical);
    if (tek && !profilBul(profiles, kod.split(/[-\s]/)[0])) {
      sonuc.lines.push({
        kind: "technical", code: tek.t.name + (tek.kartonKodu ? ` (${tek.kartonKodu})` : ""), rawCode: kod, matched: true,
        unit: birim === "adet" || birim === "kutu" ? birim : birim ? birim : "kutu", qty, note: "", confidence: guven,
        techCode: tek.t.code, kartonKodu: tek.kartonKodu,
      });
      continue;
    }
    // 2) Cam / ayna
    if (/\bAYNA\b/.test(n)) { sonuc.lines.push({ kind: "ayna", code: kod, rawCode: kod, matched: false, unit: birim || "plaka", qty, note: kod, confidence: 0.7 }); continue; }
    if (/\bCAM\b|PLAKA|MUZE|MAT CAM/.test(n) && !profilBul(profiles, kod.split(/[-\s]/)[0])) {
      sonuc.lines.push({ kind: "glass", code: kod, rawCode: kod, matched: false, unit: birim || "plaka", qty, note: kod, confidence: 0.7 });
      continue;
    }
    // 3) Çerçeve profili
    const c = cerceveKodu(kod, profiles);
    if (c.matched) {
      const u = birim === "koli" || birim === "boy" || birim === "metre" ? birim : birim === "" ? "koli" : birim;
      sonuc.lines.push({ kind: "frame", code: c.code, rawCode: kod, matched: true, unit: u, qty, note: c.note || (birim ? "" : "birim yazılmamış, koli varsayıldı"), confidence: c.note ? 0.8 : guven });
      continue;
    }
    // 4) Katalogda bulunamayan ama kod gibi görünen çerçeve satırı: atılmaz, "?" ile forma gider
    if ((birim === "koli" || birim === "boy" || birim === "metre") && /\d/.test(kod) && kod.split(" ").length <= 4) {
      sonuc.lines.push({ kind: "frame", code: c.code, rawCode: kod, matched: false, unit: birim, qty, note: "katalogda bulunamadı, kodu kontrol edin", confidence: 0.4 });
      continue;
    }
    // 5) Okunamadı → yapay zekâya / kullanıcıya
    sonuc.kalan.push(ham);
  }
  return sonuc;
}
