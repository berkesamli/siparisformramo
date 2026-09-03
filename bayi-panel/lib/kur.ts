// USD kuru — TCMB günlük kur XML'inden (ForexSelling). Erişilemezse null döner;
// bayi manuel kur girer. Sonuç gün bazında Blob'a yazılır (tekrar istek atılmaz).
import { readJson, writeJson, istanbulDateKey } from "./store";

interface KurKaydi {
  dateKey: string;
  usd: number;
  source: "tcmb";
  fetchedAt: string;
}

export async function getUsdRate(): Promise<KurKaydi | null> {
  const dateKey = istanbulDateKey();
  const cached = await readJson<KurKaydi>(`kur/${dateKey}.json`);
  if (cached && cached.usd > 0) return cached;

  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 6000);
    const res = await fetch("https://www.tcmb.gov.tr/kurlar/today.xml", {
      signal: ctrl.signal,
      cache: "no-store",
    });
    clearTimeout(t);
    if (!res.ok) return null;
    const xml = await res.text();
    const m = /<Currency[^>]*CurrencyCode="USD"[\s\S]*?<ForexSelling>([\d.]+)<\/ForexSelling>/.exec(xml);
    const usd = m ? Number(m[1]) : 0;
    if (!(usd > 0)) return null;
    const rec: KurKaydi = { dateKey, usd, source: "tcmb", fetchedAt: new Date().toISOString() };
    await writeJson(`kur/${dateKey}.json`, rec);
    return rec;
  } catch {
    return null;
  }
}
