// Mesajlar veri katmanı — Postgres (DATABASE_URL). Şema ilk kullanımda
// kendiliğinden kurulur (CREATE TABLE IF NOT EXISTS). Sunucusuz ortamda
// bağlantı havuzu küçük tutulur. Testlerde pg-mem ile aynı arayüz kullanılır
// (setDbPool).
//
// Tablolar: mesaj_konusma (konuşma başlığı, karşı taraf, durum, atama),
// mesaj (tek tek mesajlar), mesaj_senk (kanal/hesap başına son senkron).

import type { Pool as PgPool } from "pg";
import type { Ek, Kanal, Konusma, KonusmaDurum, KonusmaFiltre, Mesaj, Yon } from "./tur";

type Sorgu = { query: (text: string, params?: unknown[]) => Promise<{ rows: any[]; rowCount: number | null }> };

let havuz: Sorgu | null = null;
let kuruldu: Promise<void> | null = null;

/**
 * Postgres bağlantı adresi. Vercel'in Neon entegrasyonu değişkeni seçilen ön eke göre adlandırır
 * (DATABASE_URL, STORAGE_URL, POSTGRES_URL …); bilinen adlar sırayla, sonra "postgres://" ile
 * başlayan herhangi bir *_URL değişkeni (havuzlu olan tercih edilir) bulunur.
 */
export function dbUrl(): string {
  const bilinen = ["DATABASE_URL", "POSTGRES_URL", "STORAGE_URL", "DATABASE_POSTGRES_URL", "STORAGE_POSTGRES_URL", "NEON_DATABASE_URL"];
  for (const k of bilinen) {
    const v = (process.env[k] || "").trim();
    if (v) return v;
  }
  const adaylar = Object.entries(process.env)
    .filter(([k, v]) => /_URL$/.test(k) && /^postgres(ql)?:\/\//i.test(String(v || "").trim()))
    .sort(([a], [b]) => Number(/UNPOOLED|NON_POOLING|PRISMA|NO_SSL/.test(a)) - Number(/UNPOOLED|NON_POOLING|PRISMA|NO_SSL/.test(b)));
  return adaylar[0] ? String(adaylar[0][1]).trim() : "";
}

export function dbConfigured(): boolean {
  return Boolean(havuz || dbUrl());
}

/** Testler / farklı sürücüler için havuzu dışarıdan ver. */
export function setDbPool(p: Sorgu | null, semaHazir = false) {
  havuz = p;
  kuruldu = semaHazir ? Promise.resolve() : null;
}

async function pool(): Promise<Sorgu> {
  if (havuz) return havuz;
  const url = dbUrl();
  if (!url) throw new Error("Mesajlar için veri tabanı ayarlanmamış (DATABASE_URL).");
  const { Pool } = (await import("pg")) as { Pool: typeof PgPool };
  const ssl = /localhost|127\.0\.0\.1/.test(url) ? undefined : { rejectUnauthorized: false };
  havuz = new Pool({ connectionString: url, max: 3, idleTimeoutMillis: 10_000, ssl });
  return havuz;
}

const SEMA = `
CREATE TABLE IF NOT EXISTS mesaj_konusma (
  id TEXT PRIMARY KEY,
  kanal TEXT NOT NULL,
  hesap TEXT NOT NULL DEFAULT '',
  dis_kimlik TEXT NOT NULL,
  ad TEXT NOT NULL DEFAULT '',
  baslik TEXT NOT NULL DEFAULT '',
  musteri_id TEXT,
  musteri_tur TEXT,
  durum TEXT NOT NULL DEFAULT 'acik',
  atanan TEXT,
  son_mesaj_at TIMESTAMPTZ NOT NULL,
  son_mesaj_ozet TEXT NOT NULL DEFAULT '',
  son_gelen_at TIMESTAMPTZ,
  okunmamis INTEGER NOT NULL DEFAULT 0,
  meta JSONB NOT NULL DEFAULT '{}',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (kanal, hesap, dis_kimlik)
);
CREATE INDEX IF NOT EXISTS mesaj_konusma_son ON mesaj_konusma (son_mesaj_at DESC);
CREATE TABLE IF NOT EXISTS mesaj (
  id TEXT PRIMARY KEY,
  konusma_id TEXT NOT NULL REFERENCES mesaj_konusma(id) ON DELETE CASCADE,
  yon TEXT NOT NULL,
  govde TEXT NOT NULL DEFAULT '',
  ekler JSONB NOT NULL DEFAULT '[]',
  dis_id TEXT,
  gonderen TEXT NOT NULL DEFAULT '',
  taslak_ai BOOLEAN NOT NULL DEFAULT false,
  durum TEXT NOT NULL DEFAULT '',
  hata TEXT,
  at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS mesaj_konusma_at ON mesaj (konusma_id, at);
CREATE UNIQUE INDEX IF NOT EXISTS mesaj_dis_id ON mesaj (konusma_id, dis_id) WHERE dis_id IS NOT NULL;
CREATE TABLE IF NOT EXISTS mesaj_senk (
  anahtar TEXT PRIMARY KEY,
  deger TEXT NOT NULL DEFAULT '',
  at TIMESTAMPTZ NOT NULL DEFAULT now()
);`;

async function db(): Promise<Sorgu> {
  const p = await pool();
  if (!kuruldu) {
    kuruldu = (async () => {
      for (const s of SEMA.split(";").map((x) => x.trim()).filter(Boolean)) await p.query(s);
    })().catch((e) => { kuruldu = null; throw e; });
  }
  await kuruldu;
  return p;
}

const yeniId = (on: string) => on + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const iso = (v: unknown) => (v instanceof Date ? v.toISOString() : v ? new Date(String(v)).toISOString() : null);

function konusmaSatir(r: any): Konusma {
  return {
    id: r.id, kanal: r.kanal, hesap: r.hesap || "", disKimlik: r.dis_kimlik, ad: r.ad || "", baslik: r.baslik || "",
    musteriId: r.musteri_id || null, musteriTur: r.musteri_tur || null, durum: r.durum, atanan: r.atanan || null,
    sonMesajAt: iso(r.son_mesaj_at) || new Date(0).toISOString(), sonMesajOzet: r.son_mesaj_ozet || "",
    sonGelenAt: iso(r.son_gelen_at), okunmamis: Number(r.okunmamis) || 0,
    meta: typeof r.meta === "string" ? JSON.parse(r.meta || "{}") : r.meta || {}, createdAt: iso(r.created_at) || "",
  };
}
function mesajSatir(r: any): Mesaj {
  return {
    id: r.id, konusmaId: r.konusma_id, yon: r.yon, govde: r.govde || "",
    ekler: typeof r.ekler === "string" ? JSON.parse(r.ekler || "[]") : r.ekler || [],
    disId: r.dis_id || null, gonderen: r.gonderen || "", taslakAi: Boolean(r.taslak_ai), durum: r.durum || "",
    hata: r.hata || null, at: iso(r.at) || "",
  };
}

/** Var olan konuşmayı anahtarıyla bulur (yoksa null). */
export async function konusmaBul(kanal: Kanal, hesap: string, disKimlik: string): Promise<Konusma | null> {
  const p = await db();
  const r = await p.query("SELECT * FROM mesaj_konusma WHERE kanal = $1 AND hesap = $2 AND dis_kimlik = $3", [kanal, hesap, disKimlik]);
  return r.rows[0] ? konusmaSatir(r.rows[0]) : null;
}

/** Konuşmayı bulur, yoksa açar; ad/başlık/meta yeni bilgiyle güncellenir (boş gelen alanlar eskisini ezmez). */
export async function konusmaBulVeyaOlustur(k: {
  kanal: Kanal; hesap: string; disKimlik: string; ad?: string; baslik?: string; meta?: Record<string, unknown>;
}): Promise<Konusma> {
  const p = await db();
  const r = await p.query("SELECT * FROM mesaj_konusma WHERE kanal = $1 AND hesap = $2 AND dis_kimlik = $3", [k.kanal, k.hesap, k.disKimlik]);
  if (r.rows[0]) {
    const eski = konusmaSatir(r.rows[0]);
    const ad = k.ad || eski.ad, baslik = k.baslik || eski.baslik;
    const meta = { ...eski.meta, ...(k.meta || {}) };
    if (ad !== eski.ad || baslik !== eski.baslik || JSON.stringify(meta) !== JSON.stringify(eski.meta)) {
      await p.query("UPDATE mesaj_konusma SET ad = $2, baslik = $3, meta = $4 WHERE id = $1", [eski.id, ad, baslik, JSON.stringify(meta)]);
    }
    return { ...eski, ad, baslik, meta };
  }
  const id = yeniId("k");
  const now = new Date();
  await p.query(
    "INSERT INTO mesaj_konusma (id, kanal, hesap, dis_kimlik, ad, baslik, son_mesaj_at, meta) VALUES ($1,$2,$3,$4,$5,$6,$7,$8)",
    [id, k.kanal, k.hesap, k.disKimlik, k.ad || "", k.baslik || "", now, JSON.stringify(k.meta || {})]
  );
  return { id, kanal: k.kanal, hesap: k.hesap, disKimlik: k.disKimlik, ad: k.ad || "", baslik: k.baslik || "", musteriId: null, musteriTur: null, durum: "acik", atanan: null, sonMesajAt: now.toISOString(), sonMesajOzet: "", sonGelenAt: null, okunmamis: 0, meta: k.meta || {}, createdAt: now.toISOString() };
}

/** Mesaj ekler; aynı dis_id ikinci kez gelirse eklemez (webhook tekrarları). Konuşma özetini günceller. */
export async function mesajEkle(m: {
  konusmaId: string; yon: Yon; govde: string; ekler?: Ek[]; disId?: string | null; gonderen?: string;
  taslakAi?: boolean; durum?: string; hata?: string | null; at?: Date;
  /** Sessiz gelen: okunmamış sayacını artırmaz, kapalı konuşmayı açmaz (otomatik e-postalar). */
  sessiz?: boolean;
}): Promise<{ mesaj: Mesaj; yeni: boolean }> {
  const p = await db();
  const ayni = async () => {
    if (!m.disId) return null;
    const r = await p.query("SELECT * FROM mesaj WHERE konusma_id = $1 AND dis_id = $2", [m.konusmaId, m.disId]);
    return r.rows[0] ? mesajSatir(r.rows[0]) : null;
  };
  const eski = await ayni();
  if (eski) return { mesaj: eski, yeni: false };
  const id = yeniId("m");
  const at = m.at || new Date();
  try {
    await p.query(
      "INSERT INTO mesaj (id, konusma_id, yon, govde, ekler, dis_id, gonderen, taslak_ai, durum, hata, at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)",
      [id, m.konusmaId, m.yon, m.govde, JSON.stringify(m.ekler || []), m.disId || null, m.gonderen || "", Boolean(m.taslakAi), m.durum || "", m.hata || null, at]
    );
  } catch (e: any) {
    // Aynı dış kimlik eş zamanlı iki webhook'tan geldi (tekil indeks): ilkini döndür.
    if (e?.code === "23505" || /unique|duplicate/i.test(String(e?.message))) {
      const yine = await ayni();
      if (yine) return { mesaj: yine, yeni: false };
    }
    throw e;
  }
  const ozet = (m.govde || (m.ekler?.length ? `[${m.ekler[0].tur === "image" ? "Görsel" : "Dosya"}]` : "")).slice(0, 140);
  if (m.yon === "gelen" && m.sessiz) {
    await p.query(
      "UPDATE mesaj_konusma SET son_mesaj_at = $2, son_mesaj_ozet = $3, son_gelen_at = $2 WHERE id = $1",
      [m.konusmaId, at, ozet]
    );
  } else if (m.yon === "gelen") {
    await p.query(
      "UPDATE mesaj_konusma SET son_mesaj_at = $2, son_mesaj_ozet = $3, son_gelen_at = $2, okunmamis = okunmamis + 1, durum = CASE WHEN durum = 'kapali' THEN 'acik' ELSE durum END WHERE id = $1",
      [m.konusmaId, at, ozet]
    );
  } else {
    await p.query(
      "UPDATE mesaj_konusma SET son_mesaj_at = $2, son_mesaj_ozet = $3, durum = CASE WHEN durum = 'acik' THEN 'yanitlandi' ELSE durum END WHERE id = $1",
      [m.konusmaId, at, ozet]
    );
  }
  const r = await p.query("SELECT * FROM mesaj WHERE id = $1", [id]);
  return { mesaj: mesajSatir(r.rows[0]), yeni: true };
}

export async function mesajDurumGuncelle(konusmaId: string, disId: string, durum: string, hata?: string | null): Promise<boolean> {
  const p = await db();
  const r = await p.query("UPDATE mesaj SET durum = $3, hata = $4 WHERE konusma_id = $1 AND dis_id = $2", [konusmaId, disId, durum, hata || null]);
  return (r.rowCount || 0) > 0;
}

/** Dış kimliğe (wamid) göre durum güncelle — konuşma bilinmiyorken (teslim webhook'ları). */
export async function mesajDurumDisId(disId: string, durum: string, hata?: string | null): Promise<boolean> {
  const p = await db();
  const r = await p.query("UPDATE mesaj SET durum = $2, hata = $3 WHERE dis_id = $1 AND yon = 'giden'", [disId, durum, hata || null]);
  return (r.rowCount || 0) > 0;
}

export async function konusmalar(f: KonusmaFiltre = {}, benKullanici = ""): Promise<Konusma[]> {
  const p = await db();
  const kosul: string[] = [];
  const param: unknown[] = [];
  const ekle = (sql: string, v: unknown) => { param.push(v); kosul.push(sql.replace("?", `$${param.length}`)); };
  if (f.kanal) ekle("kanal = ?", f.kanal);
  if (f.durum) ekle("durum = ?", f.durum);
  if (f.atanan === "ben") ekle("atanan = ?", benKullanici);
  else if (f.atanan) ekle("atanan = ?", f.atanan);
  if (f.q) { param.push(`%${f.q}%`); const n = param.length; kosul.push(`(ad ILIKE $${n} OR baslik ILIKE $${n} OR dis_kimlik ILIKE $${n} OR son_mesaj_ozet ILIKE $${n})`); }
  const limit = Math.min(200, Math.max(1, f.limit || 60));
  const r = await p.query(`SELECT * FROM mesaj_konusma ${kosul.length ? "WHERE " + kosul.join(" AND ") : ""} ORDER BY son_mesaj_at DESC LIMIT ${limit}`, param);
  return r.rows.map(konusmaSatir);
}

export async function konusma(id: string): Promise<Konusma | null> {
  const p = await db();
  const r = await p.query("SELECT * FROM mesaj_konusma WHERE id = $1", [id]);
  return r.rows[0] ? konusmaSatir(r.rows[0]) : null;
}

export async function mesajlar(konusmaId: string, limit = 200): Promise<Mesaj[]> {
  const p = await db();
  const r = await p.query(`SELECT * FROM mesaj WHERE konusma_id = $1 ORDER BY at ASC LIMIT ${Math.min(500, limit)}`, [konusmaId]);
  return r.rows.map(mesajSatir);
}

export async function okunduIsaretle(konusmaId: string): Promise<void> {
  const p = await db();
  await p.query("UPDATE mesaj_konusma SET okunmamis = 0 WHERE id = $1", [konusmaId]);
}

export async function konusmaGuncelle(id: string, degisiklik: { durum?: KonusmaDurum; atanan?: string | null; musteriId?: string | null; musteriTur?: "toptan" | "perakende" | null; ad?: string; meta?: Record<string, unknown> }): Promise<Konusma | null> {
  const p = await db();
  const set: string[] = []; const param: unknown[] = [id];
  const koy = (col: string, v: unknown) => { param.push(v); set.push(`${col} = $${param.length}`); };
  if (degisiklik.durum) koy("durum", degisiklik.durum);
  if (degisiklik.atanan !== undefined) koy("atanan", degisiklik.atanan);
  if (degisiklik.musteriId !== undefined) { koy("musteri_id", degisiklik.musteriId); koy("musteri_tur", degisiklik.musteriId ? degisiklik.musteriTur || null : null); }
  if (degisiklik.ad !== undefined) koy("ad", degisiklik.ad);
  if (degisiklik.meta !== undefined) koy("meta", JSON.stringify(degisiklik.meta));
  if (set.length) await p.query(`UPDATE mesaj_konusma SET ${set.join(", ")} WHERE id = $1`, param);
  return konusma(id);
}

export async function okunmamisSayisi(): Promise<number> {
  const p = await db();
  const r = await p.query("SELECT COALESCE(SUM(okunmamis), 0) AS n FROM mesaj_konusma WHERE durum <> 'kapali'");
  return Number(r.rows[0]?.n) || 0;
}

/** Kanal senkron durumu (örn. gmail:<hesap> → son UID). */
export async function senkOku(anahtar: string): Promise<{ deger: string; at: string } | null> {
  const p = await db();
  const r = await p.query("SELECT deger, at FROM mesaj_senk WHERE anahtar = $1", [anahtar]);
  return r.rows[0] ? { deger: r.rows[0].deger, at: iso(r.rows[0].at) || "" } : null;
}
export async function senkYaz(anahtar: string, deger: string): Promise<void> {
  const p = await db();
  await p.query(
    "INSERT INTO mesaj_senk (anahtar, deger, at) VALUES ($1, $2, now()) ON CONFLICT (anahtar) DO UPDATE SET deger = EXCLUDED.deger, at = now()",
    [anahtar, deger]
  );
}
