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

/** Vercel'e yapıştırırken kalan baştaki/sondaki boşluk, satır sonu ve sarmalayan tırnakları ayıklar. */
function envDeger(v: string | undefined): string {
  let s = (v || "").trim();
  if (s.length >= 2 && ((s.startsWith('"') && s.endsWith('"')) || (s.startsWith("'") && s.endsWith("'")))) s = s.slice(1, -1).trim();
  return s;
}

const ENV_ALANLARI = ["MIKRO_API_KEY", "MIKRO_FIRMA_KODU", "MIKRO_KULLANICI", "MIKRO_SIFRE", "MIKRO_CALISMA_YILI"] as const;

/** Temizlenmek zorunda kalınan ortam değişkenleri (ayarlar kartında uyarı için). */
export function mikroAyarUyarilari(): string[] {
  return ENV_ALANLARI.filter((ad) => { const v = process.env[ad]; return v !== undefined && v !== envDeger(v); });
}

export function mikroAyar(): MikroAyar {
  const url = envDeger(process.env.MIKRO_API_URL).replace(/\/+$/, "");
  return {
    url,
    apiKey: envDeger(process.env.MIKRO_API_KEY),
    firma: envDeger(process.env.MIKRO_FIRMA_KODU),
    kullanici: envDeger(process.env.MIKRO_KULLANICI),
    // Şifre de temizlenir: sondaki görünmez satır sonu MD5'i tamamen değiştirir ("Şifre Hatalı")
    sifre: envDeger(process.env.MIKRO_SIFRE),
    yil: envDeger(process.env.MIKRO_CALISMA_YILI) || istanbulGun().slice(0, 4),
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
// Kullanıcı kodu biçimleri. Mikro kullanıcı kodları çoğunlukla büyük harflidir ("SRV", "ISTANBUL").
// "istanbul" ASCII kuralla "ISTANBUL", Türkçe kuralla "İSTANBUL" olur; ikisi de denenir.
const KULLANICI_BICIMLERI: { ad: string; uret: (k: string) => string }[] = [
  { ad: "", uret: (k) => k },
  { ad: "kullanıcı BÜYÜK HARF (ASCII)", uret: (k) => k.toUpperCase() },
  { ad: "kullanıcı BÜYÜK HARF (Türkçe)", uret: (k) => k.toLocaleUpperCase("tr-TR") },
];
// Denenecek kimlik varyantları: şifre biçimi × kullanıcı kodu biçimi.
const VARYANT_SAYISI = SIFRE_BICIMLERI.length * KULLANICI_BICIMLERI.length;
let tutanBicim = 0;

/** Ayarlar kartındaki "farklı bilgilerle dene" için tek istekte geçerli kimlik; hiçbir yere yazılmaz. */
export interface KimlikOverride { kullanici?: string; sifre?: string; firma?: string; yil?: string }

function etkinAyar(o?: KimlikOverride): MikroAyar {
  const a = mikroAyar();
  if (!o) return a;
  return {
    ...a,
    kullanici: (o.kullanici || "").trim() || a.kullanici,
    sifre: o.sifre !== undefined && o.sifre !== "" ? o.sifre : a.sifre,
    firma: (o.firma || "").trim() || a.firma,
    yil: (o.yil || "").trim() || a.yil,
  };
}

function varyantBicimleri(varyant: number) {
  const s = SIFRE_BICIMLERI[varyant % SIFRE_BICIMLERI.length] || SIFRE_BICIMLERI[0];
  const k = KULLANICI_BICIMLERI[Math.floor(varyant / SIFRE_BICIMLERI.length) % KULLANICI_BICIMLERI.length] || KULLANICI_BICIMLERI[0];
  return { s, k };
}

export function mikroKimlik(varyant = tutanBicim, o?: KimlikOverride): Record<string, string> {
  const a = etkinAyar(o);
  const { s, k } = varyantBicimleri(varyant);
  return { ApiKey: a.apiKey, CalismaYili: a.yil, FirmaKodu: a.firma, KullaniciKodu: k.uret(a.kullanici), Sifre: s.uret(a.sifre, istanbulGun()) };
}

export function sifreBicimiAdi(varyant = tutanBicim): string {
  const { s, k } = varyantBicimleri(varyant);
  return k.ad ? `${s.ad} · ${k.ad}` : s.ad;
}

export interface MikroYanit<T = unknown> {
  ok: boolean;
  status: number;
  data?: T;
  raw?: string;   // JSON çözülemediyse ham metin (kısaltılmış)
  hata?: string;
  bicim?: string; // tutan kimlik varyantının adı (başarılı yanıtta)
}

const ZAMAN_ASIMI_MS = 15000;

/** Ham metot çağrısı: POST {url}/Api/APIMethods/{metot}. "Şifre Hatalı"da sıradaki şifre biçimiyle yeniden dener. */
async function mikroPost<T = unknown>(metot: string, govde: Record<string, unknown>, bicim = tutanBicim, o?: KimlikOverride, zamanAsimi = ZAMAN_ASIMI_MS): Promise<MikroYanit<T>> {
  if (!mikroConfigured()) {
    return { ok: false, status: 0, hata: "Mikro API ayarları eksik (MIKRO_API_URL, MIKRO_API_KEY, MIKRO_FIRMA_KODU, MIKRO_KULLANICI, MIKRO_SIFRE)." };
  }
  const a = mikroAyar();
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), zamanAsimi);
  let res: Response;
  try {
    res = await fetch(`${a.url}/Api/APIMethods/${metot}`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "application/json" },
      body: JSON.stringify({ ...govde, Mikro: mikroKimlik(bicim, o) }),
      signal: ctrl.signal,
      cache: "no-store",
    });
  } catch (err) {
    clearTimeout(t);
    const m = err instanceof Error ? err.message : String(err);
    return {
      ok: false, status: 0,
      hata: /abort/i.test(m)
        ? `Mikro sunucusu ${Math.round(zamanAsimi / 1000)} saniyede yanıt vermedi (Funnel ya da Desktop API servisi kapalı olabilir).`
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
      // Aynı kimliği (örn. zaten büyük harfli kullanıcı kodu) ikinci kez denemeye gerek yok
      const denenen = new Set<string>([JSON.stringify(mikroKimlik(bicim, o))]);
      for (let i = 0; i < VARYANT_SAYISI; i++) {
        if (i === bicim) continue;
        const anahtar = JSON.stringify(mikroKimlik(i, o));
        if (denenen.has(anahtar)) continue;
        denenen.add(anahtar);
        const tekrar = await mikroPost<T>(metot, govde, i, o, zamanAsimi);
        if (tekrar.ok) { if (!o) { tutanBicim = i; console.log(`Mikro: kimlik varyantı '${sifreBicimiAdi()}' tuttu.`); } return tekrar; }
        if (!/ifre|password/i.test(tekrar.hata || "")) { hataMetni = tekrar.hata || hataMetni; break; }
      }
    }
    return { ok: false, status: zarf.kod || res.status, data: (zarf.icerik ?? data) as T, raw: data === undefined ? text.slice(0, 500) : undefined, hata: hataMetni };
  }
  if (data === undefined) return { ok: false, status: res.status, raw: text.slice(0, 500), hata: "Mikro JSON yerine metin döndürdü." };
  return { ok: true, status: res.status, data: (zarf.icerik ?? data) as T, bicim: sifreBicimiAdi(bicim) };
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

/**
 * Yanıttan satır dizisini bulur. Mikro Desktop API (17.07) SqlVeriOkuV2 için
 * Data = [{"SQLResult1":[{…satır…},…]}] döndürür; tek sorguda ilk sonuç
 * kümesi açılır. Diğer biçimler: [] | {Data:[]} | {Sonuc:[]} | {rows:[]} …
 */
function satirlariAl(data: unknown): Record<string, unknown>[] | null {
  if (Array.isArray(data)) {
    const ilk = data[0] as Record<string, unknown> | undefined;
    if (data.length === 1 && ilk && typeof ilk === "object" && !Array.isArray(ilk)) {
      const anahtar = Object.keys(ilk).find((k) => /^SQLResult\d*$/i.test(k) && Array.isArray(ilk[k]));
      if (anahtar) return ilk[anahtar] as Record<string, unknown>[];
    }
    return data as Record<string, unknown>[];
  }
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
  bicim?: string;
  status?: number; // 0: sunucuya ulaşılamadı, 401: kimlik; diğerleri SQL/uygulama hatası
  ok: boolean;
  rows: Record<string, unknown>[];
  hata?: string;
  raw?: unknown;   // ham yanıt (biçimi öğrenmek/ayıklamak için)
}

/** Modül içi: yalnızca bu dosyada yazılı sabit SELECT sorguları için. */
async function sabitSorgu(sql: string, o?: KimlikOverride, zamanAsimi?: number): Promise<SqlSonuc> {
  const r = await mikroPost("SqlVeriOkuV2", { SQLSorgu: sql }, tutanBicim, o, zamanAsimi);
  if (!r.ok) return { ok: false, rows: [], hata: r.hata, raw: r.data ?? r.raw, status: r.status };
  const rows = satirlariAl(r.data);
  if (!rows) return { ok: false, rows: [], hata: "Yanıtta satır listesi bulunamadı.", raw: r.data };
  return { ok: true, rows, raw: r.data, bicim: r.bicim };
}

/** Ham yanıtın kısaltılmış JSON'u (ekranda inceleme için). */
export function hamOzet(v: unknown, n = 1500): string {
  try { return JSON.stringify(v).slice(0, n); } catch { return String(v).slice(0, n); }
}

/** Tek tırnaklı SQL sabiti için kaçış; kod alanları kısa ve satırsız. */
const sqlStr = (s: string) => s.replace(/'/g, "''").replace(/[\r\n\t]/g, " ").slice(0, 40);
/** LIKE deseni için ek kaçış: joker karakterler ([ % _) düz metin sayılır. */
const likeStr = (s: string) => sqlStr(s).replace(/[[%_]/g, (c) => `[${c}]`);

// Kısa süreli bellek içi önbellek (sunucusuz örnek başına): aynı cari birkaç
// saniye arayla tekrar sorulduğunda Mikro'ya gidilmez.
const ONBELLEK_MS = 60_000;
const onbellek = new Map<string, { t: number; v: unknown }>();
function onbellekAl<T>(k: string): T | undefined {
  const c = onbellek.get(k);
  if (c && Date.now() - c.t < ONBELLEK_MS) return c.v as T;
  if (c) onbellek.delete(k);
  return undefined;
}
function onbellekKoy(k: string, v: unknown) {
  if (onbellek.size > 500) onbellek.clear();
  onbellek.set(k, { t: Date.now(), v });
}

// ---------- Sabit sorgular ----------

export interface CariSatir { cari_kod: string; unvan: string }

/** Bağlantı denemesi: ilk 5 cari kart. */
export async function baglantiTesti(o?: KimlikOverride): Promise<{ ok: boolean; cariler: CariSatir[]; hata?: string; ham?: string; sutunlar?: string[]; bicim?: string }> {
  const r = await sabitSorgu("SELECT TOP 5 cari_kod, cari_unvan1 FROM CARI_HESAPLAR ORDER BY cari_kod", o);
  if (!r.ok) return { ok: false, cariler: [], hata: r.hata, ham: hamOzet(r.raw) };
  const ilk = r.rows[0];
  return {
    ok: true,
    cariler: r.rows.map((x) => ({ cari_kod: String(x.cari_kod ?? ""), unvan: String(x.cari_unvan1 ?? x.unvan ?? "") })),
    sutunlar: ilk && typeof ilk === "object" ? Object.keys(ilk) : [],
    ham: hamOzet(r.raw),
    bicim: r.bicim,
  };
}

export interface CariOzet {
  cariKod: string;
  unvan: string;
  borc: number;
  alacak: number;
  bakiye: number;          // borç − alacak (pozitif: müşteri borçlu)
  sonHareket: string | null;
}

export interface CariKart { cariKod: string; unvan: string; unvan2?: string }

/**
 * Cari kartlarda ad / kod araması (en fazla 15 sonuç). Yazılan her kelime
 * ünvanda ya da kodda geçmeli; Türkçe ve ASCII büyük/küçük biçimler birlikte
 * denenir ("istanbul" → İSTANBUL ve ISTANBUL). Sorgu bu dosyada sabittir,
 * kelimeler kaçışlanarak LIKE desenine girer.
 */
export async function cariAra(sorgu: string): Promise<{ ok: boolean; cariler: CariKart[]; hata?: string }> {
  const kelimeler = sorgu.replace(/\s+/g, " ").trim().slice(0, 60).split(" ").filter((k) => k.length >= 2).slice(0, 4);
  if (!kelimeler.length) return { ok: false, cariler: [], hata: "En az iki harf yazın." };
  const anahtar = `ara:${kelimeler.join(" ").toLocaleLowerCase("tr-TR")}`;
  const hazir = onbellekAl<CariKart[]>(anahtar);
  if (hazir) return { ok: true, cariler: hazir };
  const kosul = kelimeler
    .map((k) => {
      const bicimler = Array.from(new Set([k, k.toLocaleUpperCase("tr-TR"), k.toUpperCase(), k.toLocaleLowerCase("tr-TR")]));
      return "(" + bicimler.map((b) => `cari_unvan1 LIKE '%${likeStr(b)}%' OR cari_unvan2 LIKE '%${likeStr(b)}%' OR cari_kod LIKE '${likeStr(b)}%'`).join(" OR ") + ")";
    })
    .join(" AND ");
  const r = await sabitSorgu(`SELECT TOP 15 cari_kod, cari_unvan1, cari_unvan2 FROM CARI_HESAPLAR WHERE ${kosul} ORDER BY cari_unvan1`);
  if (!r.ok) return { ok: false, cariler: [], hata: r.hata };
  const cariler = r.rows
    .map((x) => ({ cariKod: String(x.cari_kod ?? "").trim(), unvan: String(x.cari_unvan1 ?? "").trim(), unvan2: String(x.cari_unvan2 ?? "").trim() || undefined }))
    .filter((x) => x.cariKod);
  onbellekKoy(anahtar, cariler);
  return { ok: true, cariler };
}

const num = (v: unknown) => (typeof v === "number" ? v : Number(String(v ?? "").replace(",", ".")) || 0);

/** Ağ ya da kimlik hatası: sorgu biçimini değiştirmek işe yaramaz, yedek sorgu denenmez. */
const zincirDursun = (r: SqlSonuc) => r.status === 0 || /ifre|password|ulaşılamadı|yanıt vermedi/i.test(r.hata || "");

/** Ayarlar kartı için: hareket tablosundan son birkaç satır (Mikro ekstresiyle karşılaştırmak için). */
export async function hareketOrnegi(cariKod: string): Promise<{ ok: boolean; satirlar?: Record<string, unknown>[]; hata?: string }> {
  const kod = sqlStr(cariKod.trim());
  if (!kod) return { ok: false, hata: "Cari kodu boş." };
  const r = await sabitSorgu(
    `SELECT TOP 3 cha_tarihi, cha_tip, cha_meblag FROM CARI_HESAP_HAREKETLERI WHERE cha_kod = '${kod}' ORDER BY cha_tarihi DESC`
  );
  if (!r.ok) return { ok: false, hata: r.hata };
  return { ok: true, satirlar: r.rows };
}

/**
 * Cari hesabın bakiyesi: Mikro'nun CARI_HESAPLAR ve CARI_HESAP_HAREKETLERI
 * tablolarından borç (cha_tip 0) ve alacak (cha_tip 1) toplamı; bakiye = borç − alacak.
 * Vade hesabı yapılmaz; kuruşu kuruşuna Mikro ekstresi esastır.
 */
export async function cariOzet(cariKod: string): Promise<{ ok: boolean; ozet?: CariOzet; hata?: string }> {
  const kod = sqlStr(cariKod.trim());
  if (!kod) return { ok: false, hata: "Cari kodu boş." };
  const anahtar = `ozet:${kod}`;
  const hazir = onbellekAl<CariOzet>(anahtar);
  if (hazir) return { ok: true, ozet: hazir };
  const govde =
    `FROM CARI_HESAPLAR c LEFT JOIN CARI_HESAP_HAREKETLERI h ON h.cha_kod = c.cari_kod ` +
    `WHERE c.cari_kod = '${kod}' GROUP BY c.cari_kod, c.cari_unvan1`;
  const borcAlacak =
    `SELECT c.cari_kod AS cari_kod, c.cari_unvan1 AS unvan, ` +
    `ISNULL(SUM(CASE WHEN h.cha_tip = 0 THEN h.cha_meblag ELSE 0 END), 0) AS borc, ` +
    `ISNULL(SUM(CASE WHEN h.cha_tip = 1 THEN h.cha_meblag ELSE 0 END), 0) AS alacak`;
  // Son hareket tarihiyle birlikte; o sütun bu kurulumda sorun çıkarırsa yalnızca borç/alacak
  let r = await sabitSorgu(`${borcAlacak}, MAX(h.cha_tarihi) AS son_hareket ${govde}`);
  if (!r.ok && !zincirDursun(r)) r = await sabitSorgu(`${borcAlacak} ${govde}`);
  if (!r.ok) return { ok: false, hata: r.hata };
  const row = r.rows[0];
  if (!row) return { ok: false, hata: `Mikro'da '${cariKod}' kodlu cari bulunamadı.` };
  if (!("borc" in row) || !("unvan" in row)) {
    return { ok: false, hata: `Yanıt beklenen sütunları taşımıyor. Gelen: ${hamOzet(r.raw, 800)}` };
  }
  const borc = num(row.borc), alacak = num(row.alacak);
  const ozet: CariOzet = {
    cariKod: String(row.cari_kod ?? cariKod),
    unvan: String(row.unvan ?? ""),
    borc, alacak,
    bakiye: Math.round((borc - alacak) * 100) / 100,
    sonHareket: row.son_hareket ? String(row.son_hareket).slice(0, 10) : null,
  };
  onbellekKoy(anahtar, ozet);
  return { ok: true, ozet };
}

// ---------- Stok: depo bazında miktar (günlük Excel'in yerine) ----------
//
// Excel'deki "DEPO ADI / STOK İSMİ / MİKTAR" raporunun karşılığı: STOKLAR
// tablosundaki kalemler için Ankara ve İstanbul depolarındaki güncel miktar.
// Depo numaraları DEPOLAR tablosundan adına göre bulunur (adında ANKARA /
// İSTANBUL geçen depolar toplanır); MIKRO_DEPO_ANKARA / MIKRO_DEPO_ISTANBUL
// ("1,3" gibi) tanımlıysa onlar kullanılır. Miktar önce Mikro'nun kendi
// fonksiyonuyla (fn_DepodakiMiktar), o yoksa STOK_HAREKETLERI toplamıyla okunur.

export interface MikroDepo { no: number; ad: string }
export interface MikroStokSatir { kod: string; isim: string; birim: string; ankara: number; istanbul: number }
export interface MikroStokSonuc {
  ok: boolean;
  hata?: string;
  depolar: MikroDepo[];
  ankaraDepolar: number[];
  istanbulDepolar: number[];
  yontem?: string;
  satirlar: MikroStokSatir[];
  sureMs: number;
}

const STOK_ZAMAN_ASIMI_MS = 50_000;
const depoNoListesi = (v: string | undefined) =>
  envDeger(v).split(/[,; ]+/).map((x) => parseInt(x, 10)).filter((n) => Number.isInteger(n) && n >= 0);
const depoAdiNorm = (s: string) => String(s || "").toLocaleUpperCase("tr-TR").replace(/İ/g, "I");

export async function depoListesi(): Promise<{ ok: boolean; depolar: MikroDepo[]; hata?: string }> {
  const r = await sabitSorgu("SELECT dep_no, dep_adi FROM DEPOLAR ORDER BY dep_no");
  if (!r.ok) return { ok: false, depolar: [], hata: r.hata };
  const depolar = r.rows
    .map((x) => ({ no: Number(x.dep_no), ad: String(x.dep_adi ?? "").trim() }))
    .filter((d) => Number.isInteger(d.no));
  return { ok: true, depolar };
}

/** Ankara / İstanbul depo numaraları: önce ortam değişkeni, yoksa DEPOLAR'daki ad eşleşmesi. */
async function subeDepolari(): Promise<{ depolar: MikroDepo[]; ankara: number[]; istanbul: number[]; hata?: string }> {
  const envAnk = depoNoListesi(process.env.MIKRO_DEPO_ANKARA);
  const envIst = depoNoListesi(process.env.MIKRO_DEPO_ISTANBUL);
  const d = await depoListesi();
  const depolar = d.depolar;
  const ankara = envAnk.length ? envAnk : depolar.filter((x) => depoAdiNorm(x.ad).includes("ANKARA")).map((x) => x.no);
  const istanbul = envIst.length ? envIst : depolar.filter((x) => depoAdiNorm(x.ad).includes("ISTANBUL")).map((x) => x.no);
  return { depolar, ankara, istanbul, hata: d.ok ? undefined : d.hata };
}

let tutanStokYontemi = 0;

/** Depo bazında güncel miktarlar; adında PROF geçen kalemler (Excel'deki PROFİL süzgeciyle aynı). */
export async function stokMiktarlari(): Promise<MikroStokSonuc> {
  const t0 = Date.now();
  const sd = await subeDepolari();
  const bos: MikroStokSonuc = { ok: false, depolar: sd.depolar, ankaraDepolar: sd.ankara, istanbulDepolar: sd.istanbul, satirlar: [], sureMs: 0 };
  if (!sd.ankara.length && !sd.istanbul.length) {
    return { ...bos, hata: sd.hata ? `Depo listesi okunamadı: ${sd.hata}` : `Adında ANKARA ya da İSTANBUL geçen depo bulunamadı (${sd.depolar.map((d) => `${d.no}: ${d.ad}`).join(", ") || "depo yok"}). MIKRO_DEPO_ANKARA / MIKRO_DEPO_ISTANBUL ile numara verin.`, sureMs: Date.now() - t0 };
  }
  const depolar = Array.from(new Set([...sd.ankara, ...sd.istanbul]));
  // Sütun adları d<no>; depo numaraları tam sayı olarak doğrulandı
  const fonksiyonlu =
    `SELECT s.sto_kod AS kod, s.sto_isim AS isim, s.sto_birim1_ad AS birim, ` +
    depolar.map((n) => `dbo.fn_DepodakiMiktar(s.sto_kod, ${n}, GETDATE()) AS d${n}`).join(", ") +
    ` FROM STOKLAR s WHERE s.sto_isim LIKE '%PROF%'`;
  const hareketli =
    `SELECT s.sto_kod AS kod, s.sto_isim AS isim, s.sto_birim1_ad AS birim, ` +
    depolar.map((n) =>
      `ISNULL(SUM(CASE WHEN h.sth_tip = 0 AND h.sth_giris_depo_no = ${n} THEN h.sth_miktar ELSE 0 END), 0) - ` +
      `ISNULL(SUM(CASE WHEN h.sth_tip = 1 AND h.sth_cikis_depo_no = ${n} THEN h.sth_miktar ELSE 0 END), 0) AS d${n}`
    ).join(", ") +
    ` FROM STOKLAR s LEFT JOIN STOK_HAREKETLERI h ON h.sth_stok_kod = s.sto_kod WHERE s.sto_isim LIKE '%PROF%' GROUP BY s.sto_kod, s.sto_isim, s.sto_birim1_ad`;
  const yontemler = [
    { ad: "fn_DepodakiMiktar", sql: fonksiyonlu },
    { ad: "STOK_HAREKETLERI toplamı", sql: hareketli },
  ];
  const sira = tutanStokYontemi ? [tutanStokYontemi, ...yontemler.map((_, i) => i).filter((i) => i !== tutanStokYontemi)] : yontemler.map((_, i) => i);
  let r: SqlSonuc | undefined;
  let yontem = "";
  for (const i of sira) {
    r = await sabitSorgu(yontemler[i].sql, undefined, STOK_ZAMAN_ASIMI_MS);
    if (r.ok) { yontem = yontemler[i].ad; tutanStokYontemi = i; break; }
    if (zincirDursun(r)) return { ...bos, hata: r.hata, sureMs: Date.now() - t0 };
  }
  if (!r || !r.ok) return { ...bos, hata: r?.hata || "Stok sorgusu yanıt vermedi.", sureMs: Date.now() - t0 };
  const satirlar: MikroStokSatir[] = r.rows.map((x) => {
    const topla = (nolar: number[]) => nolar.reduce((s, n) => s + num(x[`d${n}`]), 0);
    return {
      kod: String(x.kod ?? "").trim(),
      isim: String(x.isim ?? "").trim(),
      birim: String(x.birim ?? "").trim(),
      ankara: Math.round(topla(sd.ankara) * 100) / 100,
      istanbul: Math.round(topla(sd.istanbul) * 100) / 100,
    };
  }).filter((x) => x.kod || x.isim);
  return { ok: true, depolar: sd.depolar, ankaraDepolar: sd.ankara, istanbulDepolar: sd.istanbul, yontem, satirlar, sureMs: Date.now() - t0 };
}
