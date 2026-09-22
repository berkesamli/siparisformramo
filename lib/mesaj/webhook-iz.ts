// Webhook izi: Meta'dan (WhatsApp / Instagram) son gelen olayın zamanı ve özeti.
// Kurulum sorunlarını ayırt etmek için: hiç olay gelmiyorsa Meta'da abonelik yok;
// olay geliyor ama imza reddediliyorsa APP_SECRET yanlış; olay var ama mesaj yoksa
// "messages" alanına abone olunmamış. Veri tabanı varsa mesaj_senk'e yazılır,
// yoksa yalnızca süreç belleğinde tutulur (soğuk başlangıçta silinir).

import { dbConfigured, senkOku, senkYaz } from "./db";

export type IzKanal = "whatsapp" | "instagram";
export interface Iz { at: string; tur: string; ozet: string }

const bellek = new Map<string, Iz>();

export async function izKaydet(kanal: IzKanal, tur: string, ozet: string): Promise<void> {
  const iz: Iz = { at: new Date().toISOString(), tur, ozet: ozet.slice(0, 200) };
  bellek.set(kanal, iz);
  if (!dbConfigured()) return;
  try { await senkYaz(`iz:${kanal}`, JSON.stringify(iz)); } catch { /* iz kaybı önemsiz */ }
}

export async function izOku(kanal: IzKanal): Promise<Iz | null> {
  if (dbConfigured()) {
    try {
      const r = await senkOku(`iz:${kanal}`);
      if (r?.deger) return JSON.parse(r.deger) as Iz;
    } catch { /* belleğe düş */ }
  }
  return bellek.get(kanal) || null;
}
