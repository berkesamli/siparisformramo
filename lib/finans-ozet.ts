// Aylık finans özeti — finans/ozet/YYYY-MM.json.
//
// Dashboard ve rapor kartları kaynak kayıtları taramak yerine yalnızca bu
// özet dosyalarını okur (ay başına tek get). Özet bir CACHE'tir: her tahsilat/
// gider yazımında imzalı delta uygulanır; bozulursa rebuildOzet o ayın kaynak
// kayıtlarını tarayıp yeniden kurar (sahiplere açık tamir ucu).

import { blobConfigured } from "./orders";
import type { Branch } from "./customers";
import {
  listTahsilatByMonths,
  type Tahsilat,
  type TahsilatYontem,
  type ParaBirimi,
} from "./tahsilat";

// Şube anahtarı: eski kayıtlarda şube olmayabilir → "belirsiz" kovası.
export type SubeKey = Branch | "belirsiz";

export interface SubeOzet {
  tahsilat: Partial<Record<TahsilatYontem, number>>;
  tahsilatToplam: number; // yalnızca TL kayıtlar (döviz ayrı izlenir)
  tahsilatDoviz: Partial<Record<Exclude<ParaBirimi, "TL">, number>>;
  gider: Record<string, number>; // kategori → toplam (Faz 2'de dolar)
  giderToplam: number;
  // Kasa kırılımı: nakit = nakit tahsilat − nakit gider tarafında hesaplanır;
  // banka = havale + kredi kartı + tahsil edilen çek (çek tahsili Faz 2'de).
  kasaGiris: { nakit: number; banka: number };
  kasaCikis: { nakit: number; banka: number };
}

export interface FinansOzet {
  month: string; // YYYY-MM
  updatedAt: string;
  sube: Partial<Record<SubeKey, SubeOzet>>;
}

const path = (month: string) => `finans/ozet/${month}.json`;

const bosSube = (): SubeOzet => ({
  tahsilat: {},
  tahsilatToplam: 0,
  tahsilatDoviz: {},
  gider: {},
  giderToplam: 0,
  kasaGiris: { nakit: 0, banka: 0 },
  kasaCikis: { nakit: 0, banka: 0 },
});

const r2 = (n: number) => Math.round(n * 100) / 100;

export async function getOzet(month: string): Promise<FinansOzet | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(path(month), { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as FinansOzet;
  } catch {
    return null;
  }
}

async function putOzet(o: FinansOzet): Promise<void> {
  const { put } = await import("@vercel/blob");
  await put(path(o.month), JSON.stringify(o), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
}

/** Kasa kanalı: nakit mi banka mı? Çek/senet kasaya girmez (tahsilinde girer). */
function kasaKanal(method: TahsilatYontem): "nakit" | "banka" | null {
  if (method === "nakit") return "nakit";
  if (method === "havale" || method === "krediKarti") return "banka";
  return null; // cek, senet, diger → kasa hareketi yok (Faz 2'de tahsil olayı)
}

function applyTahsilatToOzet(o: FinansOzet, t: Tahsilat, sign: 1 | -1): void {
  const key: SubeKey = t.branch || "belirsiz";
  const s = (o.sube[key] ||= bosSube());
  if (t.currency === "TL") {
    s.tahsilat[t.method] = r2((s.tahsilat[t.method] || 0) + sign * t.amount);
    s.tahsilatToplam = r2(s.tahsilatToplam + sign * t.amount);
    const kanal = kasaKanal(t.method);
    if (kanal) s.kasaGiris[kanal] = r2(s.kasaGiris[kanal] + sign * t.amount);
  } else {
    s.tahsilatDoviz[t.currency] = r2(
      (s.tahsilatDoviz[t.currency] || 0) + sign * t.amount
    );
  }
}

/**
 * Tek tahsilatın özet deltasını uygular. Özet dosyası yoksa oluşturur.
 * Yarış koşulu riski düşük (küçük ekip); tutarsızlık şüphesinde rebuild var.
 */
export async function applyTahsilatDelta(
  t: Tahsilat,
  sign: 1 | -1
): Promise<void> {
  if (!blobConfigured()) return;
  const month = t.dateKey.slice(0, 7);
  const o =
    (await getOzet(month)) ||
    ({ month, updatedAt: "", sube: {} } as FinansOzet);
  applyTahsilatToOzet(o, t, sign);
  o.updatedAt = new Date().toISOString();
  await putOzet(o);
}

/** Ayın özetini kaynak kayıtlardan sıfırdan kurar (tamir aracı). */
export async function rebuildOzet(month: string): Promise<FinansOzet | null> {
  if (!blobConfigured()) return null;
  const o: FinansOzet = { month, updatedAt: new Date().toISOString(), sube: {} };
  const ts = await listTahsilatByMonths([month]);
  for (const t of ts) applyTahsilatToOzet(o, t, 1);
  // Faz 2: gider ve çek-tahsil olayları da buraya eklenecek.
  await putOzet(o);
  return o;
}

/** Aralıktaki ayların özetleri (eksik aylar atlanır). */
export async function getOzetRange(months: string[]): Promise<FinansOzet[]> {
  const out: FinansOzet[] = [];
  for (const m of months) {
    const o = await getOzet(m);
    if (o) out.push(o);
  }
  return out;
}
