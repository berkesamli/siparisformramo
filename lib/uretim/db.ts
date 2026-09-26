// Üretim Takvimi veri katmanı — Postgres (DATABASE_URL), mesaj modülüyle
// AYNI bağlantı havuzu (lib/mesaj/db.ts → ortakHavuz). Şema ilk kullanımda
// kendiliğinden kurulur (yalnızca idempotent CREATE/ALTER ... IF NOT EXISTS).
//
// Tablolar:
//   uretim_is     — üretilecek işler (kaynak: online/ikas, perakende/mağaza, teklif, elle)
//   uretim_olay   — iş başına olay günlüğü (durum, taşıma, şube, not, föy, ek, senkron)
//   uretim_not    — takvim gün notları (şube başına ya da ortak)
//   uretim_teklif — mağaza müşterilerine verilen teklifler (takip tarihi takvime düşer)
//   uretim_ayar   — anahtar/değer (ikas senkron imleci, sabah listesi işareti, teklif sayacı)
//
// Gün alanları (plan_tarih, teslim_tarih, tarih) TEXT "YYYY-MM-DD" tutulur:
// DATE sütunu sürücüde yerel saate çevrilip gün kayabiliyor; metin karşılaştırması
// sözlük sırasıyla doğru çalışır. Zaman damgaları TIMESTAMPTZ, dışarı ISO metin.

import { dbConfigured as mesajDbHazir, ortakHavuz, type Sorgu } from "@/lib/mesaj/db";
import type {
  GunNotu, IsDurum, IsEk, IsKalem, IsKaynak, IsOlay, OlayTur, Sube, Teklif, TeklifDurum, UretimIs,
} from "./tur";
import { DURUM_LABELS, SUBE_LABELS, isAcik } from "./tur";

export { setDbPool } from "@/lib/mesaj/db";

export function dbConfigured(): boolean {
  return mesajDbHazir();
}

let kuruldu: Promise<void> | null = null;
let havuzOverride: Sorgu | null = null;

/** Testler için: yalnızca üretim katmanına ayrı havuz ver (mesaj katmanına dokunmadan). */
export function setUretimPool(p: Sorgu | null, semaHazir = false) {
  havuzOverride = p;
  kuruldu = semaHazir ? Promise.resolve() : null;
}

const SEMA = `
CREATE TABLE IF NOT EXISTS uretim_is (
  id TEXT PRIMARY KEY,
  kaynak TEXT NOT NULL,
  kaynak_ref TEXT,
  kaynak_key TEXT NOT NULL DEFAULT '',
  baslik TEXT NOT NULL DEFAULT '',
  musteri_ad TEXT NOT NULL DEFAULT '',
  musteri_tel TEXT NOT NULL DEFAULT '',
  musteri_adres TEXT NOT NULL DEFAULT '',
  musteri_sehir TEXT NOT NULL DEFAULT '',
  sube TEXT NOT NULL DEFAULT '',
  sube_oneri TEXT NOT NULL DEFAULT '',
  sube_oneri_neden TEXT NOT NULL DEFAULT '',
  plan_tarih TEXT,
  plan_saat TEXT NOT NULL DEFAULT '',
  plan_sira INTEGER NOT NULL DEFAULT 0,
  teslim_tarih TEXT,
  durum TEXT NOT NULL DEFAULT 'yeni',
  kalemler JSONB NOT NULL DEFAULT '[]',
  adet INTEGER NOT NULL DEFAULT 0,
  tutar NUMERIC(14,2) NOT NULL DEFAULT 0,
  notlar TEXT NOT NULL DEFAULT '',
  foy_yol TEXT,
  foy_at TIMESTAMPTZ,
  ekler JSONB NOT NULL DEFAULT '[]',
  ham JSONB NOT NULL DEFAULT '{}',
  olusturan TEXT NOT NULL DEFAULT '',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  tamamlandi_at TIMESTAMPTZ,
  kaynak_at TIMESTAMPTZ
);
CREATE UNIQUE INDEX IF NOT EXISTS uretim_is_kaynak ON uretim_is (kaynak, kaynak_ref);
CREATE INDEX IF NOT EXISTS uretim_is_plan ON uretim_is (plan_tarih, plan_sira);
CREATE INDEX IF NOT EXISTS uretim_is_durum ON uretim_is (durum);
CREATE TABLE IF NOT EXISTS uretim_olay (
  id TEXT PRIMARY KEY,
  is_id TEXT NOT NULL REFERENCES uretim_is(id) ON DELETE CASCADE,
  tur TEXT NOT NULL,
  aciklama TEXT NOT NULL DEFAULT '',
  by TEXT NOT NULL DEFAULT '',
  at TIMESTAMPTZ NOT NULL DEFAULT now(),
  meta JSONB NOT NULL DEFAULT '{}'
);
CREATE INDEX IF NOT EXISTS uretim_olay_is ON uretim_olay (is_id, at);
CREATE TABLE IF NOT EXISTS uretim_not (
  id TEXT PRIMARY KEY,
  tarih TEXT NOT NULL,
  sube TEXT NOT NULL DEFAULT '',
  metin TEXT NOT NULL DEFAULT '',
  tamam BOOLEAN NOT NULL DEFAULT false,
  by TEXT NOT NULL DEFAULT '',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS uretim_not_tarih ON uretim_not (tarih);
CREATE TABLE IF NOT EXISTS uretim_teklif (
  id TEXT PRIMARY KEY,
  no TEXT NOT NULL UNIQUE,
  musteri_ad TEXT NOT NULL DEFAULT '',
  musteri_tel TEXT NOT NULL DEFAULT '',
  musteri_eposta TEXT NOT NULL DEFAULT '',
  musteri_adres TEXT NOT NULL DEFAULT '',
  musteri_id TEXT NOT NULL DEFAULT '',
  sube TEXT NOT NULL DEFAULT '',
  kalemler JSONB NOT NULL DEFAULT '[]',
  usd_kur NUMERIC(12,4) NOT NULL DEFAULT 0,
  brut NUMERIC(14,2) NOT NULL DEFAULT 0,
  iskonto NUMERIC(14,2) NOT NULL DEFAULT 0,
  toplam NUMERIC(14,2) NOT NULL DEFAULT 0,
  gecerlilik TEXT,
  takip_tarih TEXT,
  durum TEXT NOT NULL DEFAULT 'taslak',
  notlar TEXT NOT NULL DEFAULT '',
  olusturan TEXT NOT NULL DEFAULT '',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  siparis_ref TEXT NOT NULL DEFAULT '',
  siparis_key TEXT NOT NULL DEFAULT '',
  is_id TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS uretim_teklif_takip ON uretim_teklif (takip_tarih);
CREATE TABLE IF NOT EXISTS uretim_ayar (
  anahtar TEXT PRIMARY KEY,
  deger TEXT NOT NULL DEFAULT '',
  at TIMESTAMPTZ NOT NULL DEFAULT now()
);`;

async function db(): Promise<Sorgu> {
  const p = havuzOverride || (await ortakHavuz());
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
const json = <T,>(v: unknown, vars: T): T => {
  if (v == null) return vars;
  if (typeof v === "string") { try { return JSON.parse(v) as T; } catch { return vars; } }
  return v as T;
};
const num = (v: unknown) => Math.round((Number(v) || 0) * 100) / 100;
const gun = (v: unknown): string | null => {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v);
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : null;
};

// ---- Satır → nesne ----
function isSatir(r: any): UretimIs {
  return {
    id: r.id,
    kaynak: r.kaynak as IsKaynak,
    kaynakRef: r.kaynak_ref || "",
    kaynakKey: r.kaynak_key || "",
    baslik: r.baslik || "",
    musteriAd: r.musteri_ad || "",
    musteriTel: r.musteri_tel || "",
    musteriAdres: r.musteri_adres || "",
    musteriSehir: r.musteri_sehir || "",
    sube: (r.sube || "") as Sube,
    subeOneri: (r.sube_oneri || "") as Sube,
    subeOneriNeden: r.sube_oneri_neden || "",
    planTarih: gun(r.plan_tarih),
    planSaat: r.plan_saat || "",
    planSira: Number(r.plan_sira) || 0,
    teslimTarih: gun(r.teslim_tarih),
    durum: (r.durum || "yeni") as IsDurum,
    kalemler: json<IsKalem[]>(r.kalemler, []),
    adet: Number(r.adet) || 0,
    tutar: num(r.tutar),
    notlar: r.notlar || "",
    foyYol: r.foy_yol || null,
    foyAt: iso(r.foy_at),
    ekler: json<IsEk[]>(r.ekler, []),
    ham: json<Record<string, unknown>>(r.ham, {}),
    olusturan: r.olusturan || "",
    createdAt: iso(r.created_at) || "",
    updatedAt: iso(r.updated_at) || "",
    tamamlandiAt: iso(r.tamamlandi_at),
    kaynakAt: iso(r.kaynak_at),
  };
}

function olaySatir(r: any): IsOlay {
  return {
    id: r.id, isId: r.is_id, tur: r.tur as OlayTur, aciklama: r.aciklama || "", by: r.by || "",
    at: iso(r.at) || "", meta: json<Record<string, unknown>>(r.meta, {}),
  };
}

function notSatir(r: any): GunNotu {
  return {
    id: r.id, tarih: gun(r.tarih) || "", sube: (r.sube || "") as Sube, metin: r.metin || "",
    tamam: Boolean(r.tamam), by: r.by || "", createdAt: iso(r.created_at) || "", updatedAt: iso(r.updated_at) || "",
  };
}

function teklifSatir(r: any): Teklif {
  return {
    id: r.id, no: r.no, musteriAd: r.musteri_ad || "", musteriTel: r.musteri_tel || "", musteriEposta: r.musteri_eposta || "",
    musteriAdres: r.musteri_adres || "", musteriId: r.musteri_id || "", sube: (r.sube || "") as Sube,
    kalemler: json(r.kalemler, []), usdKur: Number(r.usd_kur) || 0, brut: num(r.brut), iskonto: num(r.iskonto), toplam: num(r.toplam),
    gecerlilik: gun(r.gecerlilik), takipTarih: gun(r.takip_tarih), durum: (r.durum || "taslak") as TeklifDurum,
    notlar: r.notlar || "", olusturan: r.olusturan || "", createdAt: iso(r.created_at) || "", updatedAt: iso(r.updated_at) || "",
    siparisRef: r.siparis_ref || "", siparisKey: r.siparis_key || "", isId: r.is_id || "",
  };
}

const adetHesapla = (k: IsKalem[]) => k.reduce((t, x) => t + (Math.max(0, Math.round(Number(x.adet) || 0)) || 0), 0);

// =====================================================================
// İşler
// =====================================================================

export interface IsTaslak {
  kaynak: IsKaynak;
  kaynakRef?: string;
  kaynakKey?: string;
  baslik?: string;
  musteriAd?: string;
  musteriTel?: string;
  musteriAdres?: string;
  musteriSehir?: string;
  sube?: Sube;
  subeOneri?: Sube;
  subeOneriNeden?: string;
  planTarih?: string | null;
  planSaat?: string;
  teslimTarih?: string | null;
  durum?: IsDurum;
  kalemler?: IsKalem[];
  tutar?: number;
  notlar?: string;
  ham?: Record<string, unknown>;
  kaynakAt?: string | null;
  foyYol?: string | null;
}

/** Yeni iş açar; "olustur" olayı yazılır. Gün içi sıra: o günün en sonu. */
export async function isEkle(t: IsTaslak, by: string): Promise<UretimIs> {
  const p = await db();
  const id = yeniId("u");
  const kalemler = t.kalemler || [];
  const planTarih = t.planTarih || null;
  let sira = 0;
  if (planTarih) {
    const r = await p.query("SELECT COALESCE(MAX(plan_sira), 0) AS m FROM uretim_is WHERE plan_tarih = $1", [planTarih]);
    sira = (Number(r.rows[0]?.m) || 0) + 1;
  }
  const durum: IsDurum = t.durum || (planTarih ? "planlandi" : "yeni");
  await p.query(
    `INSERT INTO uretim_is (id, kaynak, kaynak_ref, kaynak_key, baslik, musteri_ad, musteri_tel, musteri_adres, musteri_sehir,
      sube, sube_oneri, sube_oneri_neden, plan_tarih, plan_saat, plan_sira, teslim_tarih, durum, kalemler, adet, tutar, notlar,
      ham, olusturan, kaynak_at, foy_yol, foy_at, tamamlandi_at)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27)`,
    [
      id, t.kaynak, t.kaynakRef || null, t.kaynakKey || "", t.baslik || t.musteriAd || "", t.musteriAd || "", t.musteriTel || "",
      t.musteriAdres || "", t.musteriSehir || "", t.sube || "", t.subeOneri || "", t.subeOneriNeden || "", planTarih,
      t.planSaat || "", sira, t.teslimTarih || null, durum, JSON.stringify(kalemler), adetHesapla(kalemler), num(t.tutar),
      t.notlar || "", JSON.stringify(t.ham || {}), by, t.kaynakAt ? new Date(t.kaynakAt) : null, t.foyYol || null,
      t.foyYol ? new Date() : null, durum === "tamamlandi" || durum === "teslim" ? new Date() : null,
    ]
  );
  await olayEkle(id, "olustur", `İş açıldı (${t.kaynak}${t.kaynakRef ? " · " + t.kaynakRef : ""})`, by, {});
  return (await isBul(id)) as UretimIs;
}

export async function isBul(id: string): Promise<UretimIs | null> {
  const p = await db();
  const r = await p.query("SELECT * FROM uretim_is WHERE id = $1", [id]);
  return r.rows[0] ? isSatir(r.rows[0]) : null;
}

export async function isBulKaynak(kaynak: IsKaynak, kaynakRef: string): Promise<UretimIs | null> {
  const p = await db();
  const r = await p.query("SELECT * FROM uretim_is WHERE kaynak = $1 AND kaynak_ref = $2", [kaynak, kaynakRef]);
  return r.rows[0] ? isSatir(r.rows[0]) : null;
}

export interface IsFiltre {
  from?: string;        // plan_tarih >= from
  to?: string;          // plan_tarih <= to
  plansiz?: boolean;    // plan_tarih IS NULL olanlar da gelsin (yalnızca açık olanlar)
  kaynak?: IsKaynak | "";
  sube?: Sube | "hepsi";
  durum?: IsDurum | "acik" | "kapali" | "";
  q?: string;
  limit?: number;
}

/**
 * Takvim sorgusu: aralıktaki planlı işler (+ istenirse planlanmamış açık işler).
 * Sıra: plan tarihi, gün içi sıra, oluşturma.
 */
export async function isler(f: IsFiltre = {}): Promise<UretimIs[]> {
  const p = await db();
  const kosul: string[] = [];
  const param: unknown[] = [];
  const ekle = (sql: string, v: unknown) => { param.push(v); kosul.push(sql.replace("?", `$${param.length}`)); };
  const aralik: string[] = [];
  if (f.from) { param.push(f.from); aralik.push(`plan_tarih >= $${param.length}`); }
  if (f.to) { param.push(f.to); aralik.push(`plan_tarih <= $${param.length}`); }
  if (aralik.length || f.plansiz) {
    const parca: string[] = [];
    if (aralik.length) parca.push(`(${aralik.join(" AND ")})`);
    if (f.plansiz) parca.push(`(plan_tarih IS NULL AND durum IN ('yeni','planlandi','uretimde','hazir'))`);
    kosul.push(`(${parca.join(" OR ")})`);
  }
  if (f.kaynak) ekle("kaynak = ?", f.kaynak);
  if (f.sube && f.sube !== "hepsi") ekle("sube = ?", f.sube);
  if (f.durum === "acik") kosul.push("durum IN ('yeni','planlandi','uretimde','hazir')");
  else if (f.durum === "kapali") kosul.push("durum IN ('teslim','tamamlandi','iptal')");
  else if (f.durum) ekle("durum = ?", f.durum);
  if (f.q) {
    param.push(`%${f.q}%`);
    const n = param.length;
    kosul.push(`(baslik ILIKE $${n} OR musteri_ad ILIKE $${n} OR musteri_tel ILIKE $${n} OR kaynak_ref ILIKE $${n} OR notlar ILIKE $${n} OR kalemler::text ILIKE $${n})`);
  }
  const limit = Math.min(2000, Math.max(1, f.limit || 800));
  const r = await p.query(
    `SELECT * FROM uretim_is ${kosul.length ? "WHERE " + kosul.join(" AND ") : ""} ORDER BY plan_tarih NULLS LAST, plan_sira, created_at LIMIT ${limit}`,
    param
  );
  return r.rows.map(isSatir);
}

export interface IsDegisiklik {
  baslik?: string;
  musteriAd?: string;
  musteriTel?: string;
  musteriAdres?: string;
  musteriSehir?: string;
  sube?: Sube;
  subeOneri?: Sube;
  subeOneriNeden?: string;
  planTarih?: string | null;
  planSaat?: string;
  planSira?: number;
  teslimTarih?: string | null;
  durum?: IsDurum;
  kalemler?: IsKalem[];
  tutar?: number;
  notlar?: string;
  foyYol?: string | null;
  ekler?: IsEk[];
  ham?: Record<string, unknown>;
  kaynakAt?: string | null;
}

/**
 * Kısmi güncelleme; değişen alanlar için olay yazılır (durum, taşıma, şube,
 * föy, düzenleme). Plan tarihi değişince sıra o günün sonuna alınır
 * (planSira verilmediyse). Tamamlanma/teslim zamanı ilk geçişte damgalanır.
 */
export async function isGuncelle(id: string, d: IsDegisiklik, by: string, sessiz = false): Promise<UretimIs | null> {
  const p = await db();
  const eski = await isBul(id);
  if (!eski) return null;
  const set: string[] = [];
  const param: unknown[] = [id];
  const koy = (col: string, v: unknown) => { param.push(v); set.push(`${col} = $${param.length}`); };
  const olaylar: { tur: OlayTur; aciklama: string; meta?: Record<string, unknown> }[] = [];

  if (d.baslik !== undefined) koy("baslik", d.baslik);
  if (d.musteriAd !== undefined) koy("musteri_ad", d.musteriAd);
  if (d.musteriTel !== undefined) koy("musteri_tel", d.musteriTel);
  if (d.musteriAdres !== undefined) koy("musteri_adres", d.musteriAdres);
  if (d.musteriSehir !== undefined) koy("musteri_sehir", d.musteriSehir);
  if (d.subeOneri !== undefined) koy("sube_oneri", d.subeOneri);
  if (d.subeOneriNeden !== undefined) koy("sube_oneri_neden", d.subeOneriNeden);
  if (d.planSaat !== undefined) koy("plan_saat", d.planSaat);
  if (d.notlar !== undefined) koy("notlar", d.notlar);
  if (d.tutar !== undefined) koy("tutar", num(d.tutar));
  if (d.ham !== undefined) koy("ham", JSON.stringify(d.ham));
  if (d.kaynakAt !== undefined) koy("kaynak_at", d.kaynakAt ? new Date(d.kaynakAt) : null);
  if (d.ekler !== undefined) koy("ekler", JSON.stringify(d.ekler));
  if (d.kalemler !== undefined) {
    koy("kalemler", JSON.stringify(d.kalemler));
    koy("adet", adetHesapla(d.kalemler));
  }
  if (d.sube !== undefined && d.sube !== eski.sube) {
    koy("sube", d.sube);
    olaylar.push({ tur: "sube", aciklama: d.sube ? `Şube: ${SUBE_LABELS[d.sube]}` : "Şube kaldırıldı", meta: { eski: eski.sube, yeni: d.sube } });
  }
  if (d.planTarih !== undefined && d.planTarih !== eski.planTarih) {
    koy("plan_tarih", d.planTarih);
    if (d.planSira === undefined) {
      let sira = 0;
      if (d.planTarih) {
        const r = await p.query("SELECT COALESCE(MAX(plan_sira), 0) AS m FROM uretim_is WHERE plan_tarih = $1 AND id <> $2", [d.planTarih, id]);
        sira = (Number(r.rows[0]?.m) || 0) + 1;
      }
      koy("plan_sira", sira);
    }
    olaylar.push({ tur: "tasi", aciklama: d.planTarih ? `Takvime alındı: ${d.planTarih}` : "Planlanmamışa alındı", meta: { eski: eski.planTarih, yeni: d.planTarih } });
    // Planlanmamış "yeni" iş takvime konunca kendiliğinden "planlandi" olur; takvimden çıkınca "yeni"ye döner
    if (d.durum === undefined) {
      if (d.planTarih && eski.durum === "yeni") d.durum = "planlandi";
      else if (!d.planTarih && eski.durum === "planlandi") d.durum = "yeni";
    }
  }
  if (d.planSira !== undefined) koy("plan_sira", Math.max(0, Math.round(d.planSira)));
  if (d.teslimTarih !== undefined && d.teslimTarih !== eski.teslimTarih) {
    koy("teslim_tarih", d.teslimTarih);
    olaylar.push({ tur: "duzenle", aciklama: d.teslimTarih ? `Teslim tarihi: ${d.teslimTarih}` : "Teslim tarihi kaldırıldı" });
  }
  if (d.durum !== undefined && d.durum !== eski.durum) {
    koy("durum", d.durum);
    if ((d.durum === "tamamlandi" || d.durum === "teslim") && !eski.tamamlandiAt) koy("tamamlandi_at", new Date());
    if (isAcik(d.durum) && eski.tamamlandiAt) koy("tamamlandi_at", null);
    olaylar.push({ tur: "durum", aciklama: `${DURUM_LABELS[eski.durum]} → ${DURUM_LABELS[d.durum]}`, meta: { eski: eski.durum, yeni: d.durum } });
  }
  if (d.foyYol !== undefined && d.foyYol !== eski.foyYol) {
    koy("foy_yol", d.foyYol);
    koy("foy_at", d.foyYol ? new Date() : null);
    olaylar.push({ tur: "foy", aciklama: d.foyYol ? "Üretim föyü hazır" : "Föy kaldırıldı" });
  }
  if (!set.length) return eski;
  koy("updated_at", new Date());
  await p.query(`UPDATE uretim_is SET ${set.join(", ")} WHERE id = $1`, param);
  if (!sessiz) {
    for (const o of olaylar) await olayEkle(id, o.tur, o.aciklama, by, o.meta || {});
    if (!olaylar.length) await olayEkle(id, "duzenle", "İş bilgileri güncellendi", by, {});
  }
  return isBul(id);
}

/** Bir günün iş sırasını verilen id dizisine göre yazar (sürükle-bırak). */
export async function isSirala(tarih: string, ids: string[], by: string): Promise<void> {
  const p = await db();
  for (let i = 0; i < ids.length; i++) {
    await p.query("UPDATE uretim_is SET plan_tarih = $2, plan_sira = $3, updated_at = now() WHERE id = $1", [ids[i], tarih, i + 1]);
  }
  void by;
}

export async function isSil(id: string): Promise<boolean> {
  const p = await db();
  const r = await p.query("DELETE FROM uretim_is WHERE id = $1", [id]);
  return (r.rowCount || 0) > 0;
}

// ---- Olaylar ----
export async function olayEkle(isId: string, tur: OlayTur, aciklama: string, by: string, meta: Record<string, unknown> = {}): Promise<IsOlay> {
  const p = await db();
  const id = yeniId("o");
  await p.query("INSERT INTO uretim_olay (id, is_id, tur, aciklama, by, meta) VALUES ($1,$2,$3,$4,$5,$6)", [id, isId, tur, aciklama, by, JSON.stringify(meta)]);
  const r = await p.query("SELECT * FROM uretim_olay WHERE id = $1", [id]);
  return olaySatir(r.rows[0]);
}

export async function olaylar(isId: string, limit = 200): Promise<IsOlay[]> {
  const p = await db();
  const n = Math.min(500, Math.max(1, limit));
  const r = await p.query(`SELECT * FROM uretim_olay WHERE is_id = $1 ORDER BY at DESC, id DESC LIMIT ${n}`, [isId]);
  return r.rows.map(olaySatir);
}

// =====================================================================
// Kaynaktan senkron (perakende / online): var olan iş kaynağa göre tazelenir
// =====================================================================

export interface KaynakKayit {
  kaynak: IsKaynak;
  kaynakRef: string;
  kaynakKey: string;
  kaynakAt: string;            // kaynağın updatedAt'i (ISO)
  kaynakDurum: string;         // kaynağın kendi durumu (ör. "Hazırlanıyor" ya da ikas status)
  durum: IsDurum;              // eşlenmiş üretim durumu
  musteriAd: string;
  musteriTel: string;
  musteriAdres?: string;
  musteriSehir?: string;
  sube?: Sube;
  subeOneri?: Sube;
  subeOneriNeden?: string;
  teslimTarih: string | null;
  kalemler: IsKalem[];
  tutar: number;
  notlar?: string;
  ham?: Record<string, unknown>;
  foyYol?: string | null;
}

/**
 * Kaynak kaydını işe yazar: yoksa açar (plan tarihi = teslim tarihi, yoksa
 * planlanmamış kuyruk); varsa ve kaynak daha yeniyse müşteri/kalem/tutar
 * tazelenir, durum yalnızca KAYNAĞIN durumu değiştiyse eşlenir (takvimde elle
 * ilerletilen durum, kaynak dokunulmamışken ezilmez), teslim tarihi kullanıcı
 * elle değiştirmediyse (eski kaynak değerine eşitse) güncellenir. Plan tarihi
 * hiçbir zaman kaynaktan ezilmez.
 */
export async function kaynaktanYaz(k: KaynakKayit, by: string): Promise<{ is: UretimIs; yeni: boolean; guncellendi: boolean }> {
  const p = await db();
  const mevcut = await isBulKaynak(k.kaynak, k.kaynakRef);
  const hamYeni = { ...(k.ham || {}), kaynakDurum: k.kaynakDurum, kaynakTeslim: k.teslimTarih };
  if (!mevcut) {
    try {
      const is = await isEkle({
        kaynak: k.kaynak, kaynakRef: k.kaynakRef, kaynakKey: k.kaynakKey, baslik: k.musteriAd, musteriAd: k.musteriAd,
        musteriTel: k.musteriTel, musteriAdres: k.musteriAdres || "", musteriSehir: k.musteriSehir || "", sube: k.sube || "",
        subeOneri: k.subeOneri || "", subeOneriNeden: k.subeOneriNeden || "",
        planTarih: k.teslimTarih || null, teslimTarih: k.teslimTarih, durum: k.durum, kalemler: k.kalemler, tutar: k.tutar,
        notlar: k.notlar || "", ham: hamYeni, kaynakAt: k.kaynakAt, foyYol: k.foyYol || null,
      }, by);
      return { is, yeni: true, guncellendi: false };
    } catch (e: any) {
      // Aynı kaynak eş zamanlı iki senkrondan geldi (tekil indeks) → var olanı al ve tazele
      if (!(e?.code === "23505" || /unique|duplicate/i.test(String(e?.message)))) throw e;
    }
  }
  const eski = mevcut || (await isBulKaynak(k.kaynak, k.kaynakRef));
  if (!eski) throw new Error("İş kaydı bulunamadı.");
  const eskiAt = eski.kaynakAt ? Date.parse(eski.kaynakAt) : 0;
  const yeniAt = Date.parse(k.kaynakAt) || 0;
  if (yeniAt <= eskiAt && eski.kalemler.length) return { is: eski, yeni: false, guncellendi: false };
  const d: IsDegisiklik = {
    musteriAd: k.musteriAd, musteriTel: k.musteriTel, kalemler: k.kalemler, tutar: k.tutar, kaynakAt: k.kaynakAt,
    ham: { ...eski.ham, ...hamYeni },
  };
  if (k.musteriAdres !== undefined) d.musteriAdres = k.musteriAdres;
  if (k.musteriSehir !== undefined) d.musteriSehir = k.musteriSehir;
  if (!eski.baslik || eski.baslik === eski.musteriAd) d.baslik = k.musteriAd;
  if (k.notlar !== undefined && (!eski.notlar || eski.notlar === String(eski.ham.kaynakNot || ""))) { d.notlar = k.notlar; d.ham!.kaynakNot = k.notlar; }
  if (k.foyYol && !eski.foyYol) d.foyYol = k.foyYol;
  if (k.sube && !eski.sube) d.sube = k.sube;
  const eskiKaynakDurum = String(eski.ham.kaynakDurum || "");
  if (k.kaynakDurum !== eskiKaynakDurum && k.durum !== eski.durum) d.durum = k.durum;
  const eskiKaynakTeslim = (eski.ham.kaynakTeslim as string | null | undefined) ?? null;
  if (k.teslimTarih !== eski.teslimTarih && (eski.teslimTarih === eskiKaynakTeslim || eski.teslimTarih === null)) d.teslimTarih = k.teslimTarih;
  const is = await isGuncelle(eski.id, d, by, true);
  await olayEkle(eski.id, "senk", "Kaynaktan güncellendi", by, { kaynakDurum: k.kaynakDurum });
  void p;
  return { is: is || eski, yeni: false, guncellendi: true };
}

// =====================================================================
// Gün notları
// =====================================================================

export async function notEkle(n: { tarih: string; sube: Sube; metin: string }, by: string): Promise<GunNotu> {
  const p = await db();
  const id = yeniId("n");
  await p.query("INSERT INTO uretim_not (id, tarih, sube, metin, by) VALUES ($1,$2,$3,$4,$5)", [id, n.tarih, n.sube || "", n.metin, by]);
  const r = await p.query("SELECT * FROM uretim_not WHERE id = $1", [id]);
  return notSatir(r.rows[0]);
}

export async function notGuncelle(id: string, d: { tarih?: string; sube?: Sube; metin?: string; tamam?: boolean }): Promise<GunNotu | null> {
  const p = await db();
  const set: string[] = []; const param: unknown[] = [id];
  const koy = (col: string, v: unknown) => { param.push(v); set.push(`${col} = $${param.length}`); };
  if (d.tarih !== undefined) koy("tarih", d.tarih);
  if (d.sube !== undefined) koy("sube", d.sube);
  if (d.metin !== undefined) koy("metin", d.metin);
  if (d.tamam !== undefined) koy("tamam", Boolean(d.tamam));
  if (set.length) { koy("updated_at", new Date()); await p.query(`UPDATE uretim_not SET ${set.join(", ")} WHERE id = $1`, param); }
  const r = await p.query("SELECT * FROM uretim_not WHERE id = $1", [id]);
  return r.rows[0] ? notSatir(r.rows[0]) : null;
}

export async function notSil(id: string): Promise<boolean> {
  const p = await db();
  const r = await p.query("DELETE FROM uretim_not WHERE id = $1", [id]);
  return (r.rowCount || 0) > 0;
}

export async function notlar(from: string, to: string): Promise<GunNotu[]> {
  const p = await db();
  const r = await p.query("SELECT * FROM uretim_not WHERE tarih >= $1 AND tarih <= $2 ORDER BY tarih, created_at", [from, to]);
  return r.rows.map(notSatir);
}

// =====================================================================
// Teklifler
// =====================================================================

export interface TeklifTaslak {
  musteriAd: string;
  musteriTel?: string;
  musteriEposta?: string;
  musteriAdres?: string;
  musteriId?: string;
  sube?: Sube;
  kalemler: Teklif["kalemler"];
  usdKur?: number;
  brut: number;
  iskonto?: number;
  toplam: number;
  gecerlilik?: string | null;
  takipTarih?: string | null;
  durum?: TeklifDurum;
  notlar?: string;
}

/** Yıl bazlı teklif numarası: TKL-2026-001 (uretim_ayar sayacı, çakışmasız). */
async function teklifNo(p: Sorgu): Promise<string> {
  const yil = new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" }).slice(0, 4);
  const anahtar = `teklif_sayac_${yil}`;
  for (let deneme = 0; deneme < 5; deneme++) {
    const r = await p.query("SELECT deger FROM uretim_ayar WHERE anahtar = $1", [anahtar]);
    const eski = Number(r.rows[0]?.deger) || 0;
    const yeni = eski + 1;
    let ok = false;
    if (r.rows[0]) {
      const u = await p.query("UPDATE uretim_ayar SET deger = $2, at = now() WHERE anahtar = $1 AND deger = $3", [anahtar, String(yeni), String(eski)]);
      ok = (u.rowCount || 0) > 0;
    } else {
      try { await p.query("INSERT INTO uretim_ayar (anahtar, deger) VALUES ($1, $2)", [anahtar, String(yeni)]); ok = true; } catch { ok = false; }
    }
    if (ok) return `TKL-${yil}-${String(yeni).padStart(3, "0")}`;
  }
  return `TKL-${yil}-${Date.now().toString(36).toUpperCase()}`;
}

export async function teklifEkle(t: TeklifTaslak, by: string): Promise<Teklif> {
  const p = await db();
  const id = yeniId("t");
  const no = await teklifNo(p);
  await p.query(
    `INSERT INTO uretim_teklif (id, no, musteri_ad, musteri_tel, musteri_eposta, musteri_adres, musteri_id, sube, kalemler, usd_kur, brut, iskonto, toplam,
      gecerlilik, takip_tarih, durum, notlar, olusturan) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18)`,
    [id, no, t.musteriAd, t.musteriTel || "", t.musteriEposta || "", t.musteriAdres || "", t.musteriId || "", t.sube || "",
      JSON.stringify(t.kalemler || []), Number(t.usdKur) || 0, num(t.brut), num(t.iskonto), num(t.toplam), t.gecerlilik || null,
      t.takipTarih || null, t.durum || "taslak", t.notlar || "", by]
  );
  return (await teklifBul(id)) as Teklif;
}

export async function teklifBul(id: string): Promise<Teklif | null> {
  const p = await db();
  const r = await p.query("SELECT * FROM uretim_teklif WHERE id = $1", [id]);
  return r.rows[0] ? teklifSatir(r.rows[0]) : null;
}

export interface TeklifFiltre {
  durum?: TeklifDurum | "acik" | "";
  q?: string;
  takipFrom?: string;
  takipTo?: string;
  limit?: number;
}

export async function teklifler(f: TeklifFiltre = {}): Promise<Teklif[]> {
  const p = await db();
  const kosul: string[] = []; const param: unknown[] = [];
  const ekle = (sql: string, v: unknown) => { param.push(v); kosul.push(sql.replace("?", `$${param.length}`)); };
  if (f.durum === "acik") kosul.push("durum IN ('taslak','gonderildi')");
  else if (f.durum) ekle("durum = ?", f.durum);
  if (f.takipFrom) ekle("takip_tarih >= ?", f.takipFrom);
  if (f.takipTo) ekle("takip_tarih <= ?", f.takipTo);
  if (f.q) { param.push(`%${f.q}%`); const n = param.length; kosul.push(`(musteri_ad ILIKE $${n} OR musteri_tel ILIKE $${n} OR no ILIKE $${n} OR notlar ILIKE $${n})`); }
  const limit = Math.min(1000, Math.max(1, f.limit || 300));
  const r = await p.query(`SELECT * FROM uretim_teklif ${kosul.length ? "WHERE " + kosul.join(" AND ") : ""} ORDER BY created_at DESC LIMIT ${limit}`, param);
  return r.rows.map(teklifSatir);
}

export async function teklifGuncelle(id: string, d: Partial<Omit<Teklif, "id" | "no" | "createdAt" | "updatedAt" | "olusturan">>): Promise<Teklif | null> {
  const p = await db();
  const set: string[] = []; const param: unknown[] = [id];
  const koy = (col: string, v: unknown) => { param.push(v); set.push(`${col} = $${param.length}`); };
  if (d.musteriAd !== undefined) koy("musteri_ad", d.musteriAd);
  if (d.musteriTel !== undefined) koy("musteri_tel", d.musteriTel);
  if (d.musteriEposta !== undefined) koy("musteri_eposta", d.musteriEposta);
  if (d.musteriAdres !== undefined) koy("musteri_adres", d.musteriAdres);
  if (d.musteriId !== undefined) koy("musteri_id", d.musteriId);
  if (d.sube !== undefined) koy("sube", d.sube);
  if (d.kalemler !== undefined) koy("kalemler", JSON.stringify(d.kalemler));
  if (d.usdKur !== undefined) koy("usd_kur", Number(d.usdKur) || 0);
  if (d.brut !== undefined) koy("brut", num(d.brut));
  if (d.iskonto !== undefined) koy("iskonto", num(d.iskonto));
  if (d.toplam !== undefined) koy("toplam", num(d.toplam));
  if (d.gecerlilik !== undefined) koy("gecerlilik", d.gecerlilik);
  if (d.takipTarih !== undefined) koy("takip_tarih", d.takipTarih);
  if (d.durum !== undefined) koy("durum", d.durum);
  if (d.notlar !== undefined) koy("notlar", d.notlar);
  if (d.siparisRef !== undefined) koy("siparis_ref", d.siparisRef);
  if (d.siparisKey !== undefined) koy("siparis_key", d.siparisKey);
  if (d.isId !== undefined) koy("is_id", d.isId);
  if (set.length) { koy("updated_at", new Date()); await p.query(`UPDATE uretim_teklif SET ${set.join(", ")} WHERE id = $1`, param); }
  return teklifBul(id);
}

export async function teklifSil(id: string): Promise<boolean> {
  const p = await db();
  const r = await p.query("DELETE FROM uretim_teklif WHERE id = $1", [id]);
  return (r.rowCount || 0) > 0;
}

// =====================================================================
// Ayarlar (anahtar/değer) — senkron imleci, sabah listesi işareti
// =====================================================================

export async function ayarOku(anahtar: string): Promise<{ deger: string; at: string } | null> {
  const p = await db();
  const r = await p.query("SELECT deger, at FROM uretim_ayar WHERE anahtar = $1", [anahtar]);
  return r.rows[0] ? { deger: r.rows[0].deger, at: iso(r.rows[0].at) || "" } : null;
}

export async function ayarYaz(anahtar: string, deger: string): Promise<void> {
  const p = await db();
  await p.query(
    "INSERT INTO uretim_ayar (anahtar, deger, at) VALUES ($1, $2, now()) ON CONFLICT (anahtar) DO UPDATE SET deger = EXCLUDED.deger, at = now()",
    [anahtar, deger]
  );
}

// =====================================================================
// Özet ve rapor
// =====================================================================

export interface UretimOzet {
  bugun: string;
  bugunPlanli: { ankara: number; istanbul: number; atanmamis: number; adet: number };
  plansiz: number;
  geciken: number;          // plan tarihi geçmiş, hâlâ açık
  yeniOnline: number;       // kaynak online, durum yeni (henüz planlanmamış)
  teklifTakip: number;      // takip tarihi bugün ya da geçmiş, açık teklif
  haftaPlanli: number;      // bugünden 7 gün
}

export async function ozet(bugun: string): Promise<UretimOzet> {
  const p = await db();
  // Açık işler tek sorguda okunur, sayımlar JS'te yapılır (açık iş sayısı küçüktür; altı ayrı COUNT yerine tek tur)
  const r = await p.query("SELECT plan_tarih, sube, durum, kaynak, adet FROM uretim_is WHERE durum IN ('yeni','planlandi','uretimde','hazir')");
  const bugunPlanli = { ankara: 0, istanbul: 0, atanmamis: 0, adet: 0 };
  let plansiz = 0, geciken = 0, yeniOnline = 0, haftaPlanli = 0;
  const son = new Date(bugun + "T12:00:00Z"); son.setUTCDate(son.getUTCDate() + 6);
  const haftaSon = son.toISOString().slice(0, 10);
  for (const row of r.rows) {
    const t = gun(row.plan_tarih);
    const adet = Number(row.adet) || 0;
    if (!t) plansiz++;
    else if (t === bugun) {
      if (row.sube === "ankara") bugunPlanli.ankara++; else if (row.sube === "istanbul") bugunPlanli.istanbul++; else bugunPlanli.atanmamis++;
      bugunPlanli.adet += adet;
    } else if (t < bugun) geciken++;
    if (t && t >= bugun && t <= haftaSon) haftaPlanli++;
    if (row.kaynak === "online" && row.durum === "yeni") yeniOnline++;
  }
  const e = await p.query(`SELECT COUNT(*)::int AS n FROM uretim_teklif WHERE takip_tarih IS NOT NULL AND takip_tarih <= $1 AND durum IN ('taslak','gonderildi')`, [bugun]);
  return { bugun, bugunPlanli, plansiz, geciken, yeniOnline, teklifTakip: Number(e.rows[0]?.n) || 0, haftaPlanli };
}

export interface RaporGun { tarih: string; ankara: number; istanbul: number; atanmamis: number; ankaraAdet: number; istanbulAdet: number; atanmamisAdet: number }
export interface UretimRapor {
  from: string; to: string;
  gunler: RaporGun[];
  toplam: { ankara: number; istanbul: number; atanmamis: number; adet: number };
  tamamlanan: { ankara: number; istanbul: number; ortSureGun: number | null };   // aralıkta tamamlanan işler ve plan→tamam ortalama gün
  geciken: UretimIs[];
  kaynaklar: Record<string, number>;
}

/** Şube yükü raporu: aralıktaki planlı işler gün/şube kırılımı + geciken + tamamlanma süreleri. */
export async function rapor(from: string, to: string, bugun: string): Promise<UretimRapor> {
  const p = await db();
  const r = await p.query(
    `SELECT plan_tarih, sube, COUNT(*)::int AS n, COALESCE(SUM(adet),0)::int AS adet FROM uretim_is
     WHERE plan_tarih >= $1 AND plan_tarih <= $2 AND durum <> 'iptal' GROUP BY plan_tarih, sube ORDER BY plan_tarih`, [from, to]);
  const harita = new Map<string, RaporGun>();
  const bos = (t: string): RaporGun => ({ tarih: t, ankara: 0, istanbul: 0, atanmamis: 0, ankaraAdet: 0, istanbulAdet: 0, atanmamisAdet: 0 });
  for (const row of r.rows) {
    const t = gun(row.plan_tarih) || ""; if (!t) continue;
    const g = harita.get(t) || bos(t);
    const n = Number(row.n) || 0, ad = Number(row.adet) || 0;
    if (row.sube === "ankara") { g.ankara += n; g.ankaraAdet += ad; }
    else if (row.sube === "istanbul") { g.istanbul += n; g.istanbulAdet += ad; }
    else { g.atanmamis += n; g.atanmamisAdet += ad; }
    harita.set(t, g);
  }
  const gunler = [...harita.values()].sort((a, b) => a.tarih.localeCompare(b.tarih));
  const toplam = { ankara: 0, istanbul: 0, atanmamis: 0, adet: 0 };
  for (const g of gunler) { toplam.ankara += g.ankara; toplam.istanbul += g.istanbul; toplam.atanmamis += g.atanmamis; toplam.adet += g.ankaraAdet + g.istanbulAdet + g.atanmamisAdet; }
  const t = await p.query(
    `SELECT sube, plan_tarih, tamamlandi_at, created_at FROM uretim_is WHERE tamamlandi_at IS NOT NULL AND tamamlandi_at >= $1 AND tamamlandi_at < $2`,
    [new Date(from + "T00:00:00+03:00"), new Date(new Date(to + "T00:00:00+03:00").getTime() + 86_400_000)]);
  const tamamlanan = { ankara: 0, istanbul: 0, ortSureGun: null as number | null };
  let sureToplam = 0, sureN = 0;
  for (const row of t.rows) {
    if (row.sube === "ankara") tamamlanan.ankara++; else if (row.sube === "istanbul") tamamlanan.istanbul++;
    const baslangic = row.plan_tarih ? Date.parse(String(gun(row.plan_tarih)) + "T00:00:00+03:00") : Date.parse(String(iso(row.created_at)));
    const bitis = Date.parse(String(iso(row.tamamlandi_at)));
    if (baslangic && bitis && bitis >= baslangic) { sureToplam += (bitis - baslangic) / 86_400_000; sureN++; }
  }
  if (sureN) tamamlanan.ortSureGun = Math.round((sureToplam / sureN) * 10) / 10;
  const g = await p.query(`SELECT * FROM uretim_is WHERE plan_tarih IS NOT NULL AND plan_tarih < $1 AND durum IN ('yeni','planlandi','uretimde','hazir') ORDER BY plan_tarih LIMIT 200`, [bugun]);
  const k = await p.query(`SELECT kaynak, COUNT(*)::int AS n FROM uretim_is WHERE plan_tarih >= $1 AND plan_tarih <= $2 AND durum <> 'iptal' GROUP BY kaynak`, [from, to]);
  const kaynaklar: Record<string, number> = {};
  for (const row of k.rows) kaynaklar[row.kaynak] = Number(row.n) || 0;
  return { from, to, gunler, toplam, tamamlanan, geciken: g.rows.map(isSatir), kaynaklar };
}

/** Bağlantı + tablo sayıları (ayarlar / sağlık). */
export async function uretimSaglik(): Promise<{ ok: boolean; is: number; teklif: number; hata?: string }> {
  if (!dbConfigured()) return { ok: false, is: 0, teklif: 0, hata: "DATABASE_URL tanımlı değil." };
  try {
    const p = await db();
    const a = await p.query("SELECT COUNT(*)::int AS n FROM uretim_is");
    const b = await p.query("SELECT COUNT(*)::int AS n FROM uretim_teklif");
    return { ok: true, is: Number(a.rows[0]?.n) || 0, teklif: Number(b.rows[0]?.n) || 0 };
  } catch (e) {
    return { ok: false, is: 0, teklif: 0, hata: (e as Error)?.message || String(e) };
  }
}
