// Mikro Jump 17 "Desktop API" istemcisi (yalnızca OKUMA, sabit sorgular).
//
// API, Ankara'daki sunucuda Windows servisi olarak çalışır (port 8094) ve
// Tailscale Funnel ile https://…ts.net adresinden erişilir. Her istek POST +
// JSON'dur ve gövdede "Mikro" kimlik nesnesi taşır; şifre, günün tarihiyle
// birlikte MD5'lenir: md5("YYYY-MM-DD <şifre>"). Belge: apidocs.mikro.com.tr
//
// Dışarıdan serbest SQL alınmaz: bu modül yalnızca burada yazılı, parametreleri
// kaçışlanmış SELECT sorgularını çalıştırır.
//
// Ortam değişkenleri:
//   MIKRO_API_URL        https://olgaserver.tailbbb7b8.ts.net  (sonda / olmadan)
//   MIKRO_API_KEY        Mikro'nun verdiği API anahtarı
//   MIKRO_FIRMA_KODU     Giriş ekranındaki "Veri tabanı" (örn. 001)
//   MIKRO_KULLANICI      API için Mikro kullanıcı kodu
//   MIKRO_SIFRE          O kullanıcının şifresi (düz metin; MD5'i biz alırız)
//   MIKRO_CALISMA_YILI   Çalışma yılı (örn. 2026); boşsa İstanbul yılı

import { createHash } from "node:crypto";

export interface MikroAyar {
  url: string;
  apiKey: string;
  firma: string;
  kullanici: string;
  sifre: string;
  yil: string;
}

export function mikroAyar(): MikroAyar {
  const url = (process.env.MIKRO_API_URL || "").trim().replace(/\/+$/, "");
  return {
    url,
    apiKey: (process.env.MIKRO_API_KEY || "").trim(),
    firma: (process.env.MIKRO_FIRMA_KODU || "").trim(),
    kullanici: (process.env.MIKRO_KULLANICI || "").trim(),
    sifre: process.env.MIKRO_SIFRE || "",
    yil: (process.env.MIKRO_CALISMA_YILI || "").trim() || istanbulGun().slice(0, 4),
  };
}

export function mikroConfigured(): boolean {
  const a = mikroAyar();
  return Boolean(a.url && a.apiKey && a.firma && a.kullanici && a.sifre);
}

/** İstanbul günü, YYYY-MM-DD; offsetGun ile dün/yarın (gece yarısı saat farkı için). */
export function istanbulGun(offsetGun = 0): string {
  const d = new Date(Date.now() + offsetGun * 86400000);
  return d.toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" }); // en-CA → YYYY-MM-DD
}

const md5 = (s: string) => createHash("md5").update(s, "utf8").digest("hex");

/**
 * Şifre alanı biçimleri. Belgede "YYYY-MM-DD <şifre> → MD5" yazar; sahada
 * sürüme göre farklı biçimler görüldüğünden "Şifre Hatalı" gelirse sıradaki
 * denenir ve tutan biçim süreç ömrünce hatırlanır.
 */
const SIFRE_BICIMLERI: { ad: string; uret: (sifre: string, gun: string) => string }[] = [
  { ad: "md5(gün boşluk şifre)", uret: (p, g) => md5(`${g} ${p}`) },
  { ad: "md5(dün boşluk şifre)", uret: (p, g) => md5(`${istanbulGun(-1)} ${p}`) },
  { ad: "md5(gün+şifre)", uret: (p, g) => md5(`${g}${p}`) },
  { ad: "md5(şifre)", uret: (p) => md5(p) },
  { ad: "düz şifre", uret: (p) => p },
  { ad: "MD5 büyük harf", uret: (p, g) => md5(`${g} ${p}`).toUpperCase() },
];
// Denenecek kimlik varyantları: şifre biçimi × kullanıcı kodu (yazıldığı gibi / BÜYÜK HARF).
// Mikro kullanıcı kodları çoğunlukla büyük harflidir ("SRV").
const VARYANT_SAYISI = SIFRE_BICIMLERI.length * 2;
let tutanBicim = 0;

export function mikroKimlik(varyant = tutanBicim): Record<string, string> {
  const a = mikroAyar();
  const b = SIFRE_BICIMLERI[varyant % SIFRE_BICIMLERI.length] || SIFRE_BICIMLERI[0];
  const kullanici = varyant >= SIFRE_BICIMLERI.length ? a.kullanici.toLocaleUpperCase("tr-TR") : a.kullanici;
  return { ApiKey: a.apiKey, CalismaYili: a.yil, FirmaKodu: a.firma, KullaniciKodu: kullanici, Sifre: b.uret(a.sifre, istanbulGun()) };
}

export function sifreBicimiAdi(): string {
  const b = SIFRE_BICIMLERI[tutanBicim % SIFRE_BICIMLERI.length]?.ad || "?";
  return tutanBicim >= SIFRE_BICIMLERI.length ? `${b} · kullanıcı BÜYÜK HARF` : b;
}

export interface MikroYanit<T = unknown> {
  ok: boolean;
  status: number;
  data?: T;
  raw?: string;   // JSON çözülemediyse ham metin (kısaltılmış)
  hata?: string;
}

const ZAMAN_ASIMI_MS = 15000;

/** Ham metot çağrısı: POST {url}/Api/APIMethods/{metot}. "Şifre Hatalı"da sıradaki şifre biçimiyle yeniden dener. */
async function mikroPost<T = unknown>(metot: string, govde: Record<string, unknown>, bicim = tutanBicim): Promise<MikroYanit<T>> {
  if (!mikroConfigured()) {
    return { ok: false, status: 0, hata: "Mikro API ayarları eksik (MIKRO_API_URL, MIKRO_API_KEY, MIKRO_FIRMA_KODU, MIKRO_KULLANICI, MIKRO_SIFRE)." };
  }
  const a = mikroAyar();
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ZAMAN_ASIMI_MS);
  let res: Response;
  try {
    res = await fetch(`${a.url}/Api/APIMethods/${metot}`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "application/json" },
      body: JSON.stringify({ ...govde, Mikro: mikroKimlik(bicim) }),
      signal: ctrl.signal,
      cache: "no-store",
    });
  } catch (err) {
    clearTimeout(t);
    const m = err instanceof Error ? err.message : String(err);
    return {
      ok: false, status: 0,
      hata: /abort/i.test(m)
        ? "Mikro sunucusu 15 saniyede yanıt vermedi (Funnel ya da Desktop API servisi kapalı olabilir)."
        : `Mikro sunucusuna ulaşılamadı: ${m}`,
    };
  }
  clearTimeout(t);
  const text = await res.text();
  let data: unknown = undefined;
  try { data = JSON.parse(text); } catch { /* ham metin */ }

  // Zarf: {"result":[{"StatusCode":200,"Data":…,"ErrorMessage":null,"IsError":false}]}
  const zarf = zarfCoz(data);
  let hataMetni = zarf.hata || mikroHataMetni(zarf.icerik) || (!res.ok ? `HTTP ${res.status}` : "");
  if (hataMetni) {
    // Şifre biçimi tutmadıysa sıradakini dene (yalnızca ilk turda ve şifre hatasında)
    if (bicim === tutanBicim && /ifre|password/i.test(hataMetni)) {
      for (let i = 0; i < VARYANT_SAYISI; i++) {
        if (i === bicim) continue;
        const tekrar = await mikroPost<T>(metot, govde, i);
        if (tekrar.ok) { tutanBicim = i; console.log(`Mikro: kimlik varyantı '${sifreBicimiAdi()}' tuttu.`); return tekrar; }
        if (!/ifre|password/i.test(tekrar.hata || "")) { hataMetni = tekrar.hata || hataMetni; break; }
      }
    }
    return { ok: false, status: zarf.kod || res.status, data: (zarf.icerik ?? data) as T, raw: data === undefined ? text.slice(0, 500) : undefined, hata: hataMetni };
  }
  if (data === undefined) return { ok: false, status: res.status, raw: text.slice(0, 500), hata: "Mikro JSON yerine metin döndürdü." };
  return { ok: true, status: res.status, data: (zarf.icerik ?? data) as T };
}

/** Mikro Desktop API zarfını açar; zarf yoksa veriyi olduğu gibi döndürür. */
function zarfCoz(data: unknown): { icerik: unknown; hata?: string; kod?: number } {
  if (!data || typeof data !== "object" || Array.isArray(data)) return { icerik: data };
  const o = data as Record<string, unknown>;
  const dizi = Array.isArray(o.result) ? o.result : Array.isArray(o.Result) ? o.Result : null;
  const z = dizi && dizi.length ? (dizi[0] as Record<string, unknown>) : null;
  if (!z || typeof z !== "object" || !("IsError" in z || "StatusCode" in z || "Data" in z)) return { icerik: data };
  const kod = typeof z.StatusCode === "number" ? z.StatusCode : undefined;
  if (z.IsError === true || (kod && kod >= 400)) {
    const msg = typeof z.ErrorMessage === "string" && z.ErrorMessage.trim() ? z.ErrorMessage.trim() : `Mikro hata kodu ${kod ?? "?"}`;
    return { icerik: z.Data, hata: msg, kod };
  }
  let icerik: unknown = z.Data;
  if (typeof icerik === "string") { try { icerik = JSON.parse(icerik); } catch { /* düz metin */ } }
  return { icerik, kod };
}

function mikroHataMetni(data: unknown): string {
  if (!data || typeof data !== "object" || Array.isArray(data)) return "";
  const o = data as Record<string, unknown>;
  const basarili = o.Basarili === true || o.success === true || o.Sonuc === true;
  for (const k of ["error", "Error", "hata", "Hata", "HataMesaji", "Mesaj", "message", "Message"]) {
    const v = o[k];
    if (typeof v === "string" && v.trim()) {
      if ((k === "Mesaj" || k === "message" || k === "Message") && basarili) continue;
      return v.trim();
    }
  }
  if (o.Basarili === false || o.success === false || o.Sonuc === false) return "Mikro isteği reddetti.";
  return "";
}

/** Yanıttan satır dizisini bulur: [] | {Data:[]} | {data:[]} | {Sonuc:[]} | {Result:[]} | {rows:[]} … */
function satirlariAl(data: unknown): Record<string, unknown>[] | null {
  if (Array.isArray(data)) return data as Record<string, unknown>[];
  if (data && typeof data === "object") {
    const o = data as Record<string, unknown>;
    for (const k of ["Data", "data", "Sonuc", "sonuc", "Result", "result", "rows", "Rows", "Kayitlar", "Liste", "value"]) {
      if (Array.isArray(o[k])) return o[k] as Record<string, unknown>[];
    }
    const diziler = Object.values(o).filter(Array.isArray);
    if (diziler.length === 1) return diziler[0] as Record<string, unknown>[];
  }
  return null;
}

interface SqlSonuc {
  ok: boolean;
  rows: Record<string, unknown>[];
  hata?: string;
  raw?: unknown;   // ham yanıt (biçimi öğrenmek/ayıklamak için)
}

/** Modül içi: yalnızca bu dosyada yazılı sabit SELECT sorguları için. */
async function sabitSorgu(sql: string): Promise<SqlSonuc> {
  const r = await mikroPost("SqlVeriOkuV2", { SQLSorgu: sql });
  if (!r.ok) return { ok: false, rows: [], hata: r.hata, raw: r.data ?? r.raw };
  const rows = satirlariAl(r.data);
  if (!rows) return { ok: false, rows: [], hata: "Yanıtta satır listesi bulunamadı.", raw: r.data };
  return { ok: true, rows, raw: r.data };
}

/** Ham yanıtın kısaltılmış JSON'u (ekranda inceleme için). */
export function hamOzet(v: unknown, n = 1500): string {
  try { return JSON.stringify(v).slice(0, n); } catch { return String(v).slice(0, n); }
}

/** Tek tırnaklı SQL sabiti için kaçış; kod alanları kısa ve satırsız. */
const sqlStr = (s: string) => s.replace(/'/g, "''").replace(/[\r\n\t]/g, " ").slice(0, 40);

// ---------- Sabit sorgular ----------

export interface CariSatir { cari_kod: string; unvan: string }

/** Bağlantı denemesi: ilk 5 cari kart. */
export async function baglantiTesti(): Promise<{ ok: boolean; cariler: CariSatir[]; hata?: string; ham?: string; sutunlar?: string[] }> {
  const r = await sabitSorgu("SELECT TOP 5 cari_kod, cari_unvan1 FROM CARI_HESAPLAR ORDER BY cari_kod");
  if (!r.ok) return { ok: false, cariler: [], hata: r.hata, ham: hamOzet(r.raw) };
  const ilk = r.rows[0];
  return {
    ok: true,
    cariler: r.rows.map((x) => ({ cari_kod: String(x.cari_kod ?? ""), unvan: String(x.cari_unvan1 ?? x.unvan ?? "") })),
    sutunlar: ilk && typeof ilk === "object" ? Object.keys(ilk) : [],
    ham: hamOzet(r.raw),
  };
}

export interface CariOzet {
  cariKod: string;
  unvan: string;
  borc: number;
  alacak: number;
  bakiye: number;          // borç − alacak (pozitif: müşteri borçlu)
  vadesiGecen: number;     // yaklaşık: vadesi geçmiş borç − toplam alacak (0'dan küçük olamaz)
  sonHareket: string | null;
}

const num = (v: unknown) => (typeof v === "number" ? v : Number(String(v ?? "").replace(",", ".")) || 0);

/**
 * Cari hesabın bakiyesi ve vadesi geçen tutarı. Mikro'nun CARI_HESAPLAR ve
 * CARI_HESAP_HAREKETLERI tablolarından okunur (cha_tip 0 = borç, 1 = alacak).
 * Vadesi geçen tutar FIFO kapatma yapılmadan hesaplandığı için yaklaşıktır.
 */
export async function cariOzet(cariKod: string): Promise<{ ok: boolean; ozet?: CariOzet; hata?: string }> {
  const kod = sqlStr(cariKod.trim());
  if (!kod) return { ok: false, hata: "Cari kodu boş." };
  const sql =
    `SELECT c.cari_kod AS cari_kod, c.cari_unvan1 AS unvan, ` +
    `ISNULL(SUM(CASE WHEN h.cha_tip = 0 THEN h.cha_meblag ELSE 0 END), 0) AS borc, ` +
    `ISNULL(SUM(CASE WHEN h.cha_tip = 1 THEN h.cha_meblag ELSE 0 END), 0) AS alacak, ` +
    `ISNULL(SUM(CASE WHEN h.cha_tip = 0 AND h.cha_vade < GETDATE() THEN h.cha_meblag ELSE 0 END), 0) AS vadesi_gecen_borc, ` +
    `MAX(h.cha_tarihi) AS son_hareket ` +
    `FROM CARI_HESAPLAR c LEFT JOIN CARI_HESAP_HAREKETLERI h ON h.cha_kod = c.cari_kod ` +
    `WHERE c.cari_kod = '${kod}' GROUP BY c.cari_kod, c.cari_unvan1`;
  const r = await sabitSorgu(sql);
  if (!r.ok) return { ok: false, hata: r.hata };
  const row = r.rows[0];
  if (!row) return { ok: false, hata: `Mikro'da '${cariKod}' kodlu cari bulunamadı.` };
  if (!("borc" in row) || !("unvan" in row)) {
    return { ok: false, hata: `Yanıt beklenen sütunları taşımıyor. Gelen: ${hamOzet(r.raw, 800)}` };
  }
  const borc = num(row.borc), alacak = num(row.alacak);
  const vg = num(row.vadesi_gecen_borc) - alacak;
  return {
    ok: true,
    ozet: {
      cariKod: String(row.cari_kod ?? cariKod),
      unvan: String(row.unvan ?? ""),
      borc, alacak,
      bakiye: Math.round((borc - alacak) * 100) / 100,
      vadesiGecen: Math.max(0, Math.round(vg * 100) / 100),
      sonHareket: row.son_hareket ? String(row.son_hareket).slice(0, 10) : null,
    },
  };
}
