// Webhook izi: Meta'dan (WhatsApp / Instagram) son gelen olayın zamanı ve özeti.
// Kurulum sorunlarını ayırt etmek için: hiç olay gelmiyorsa Meta'da abonelik yok;
// olay geliyor ama imza reddediliyorsa APP_SECRET yanlış; olay var ama mesaj yoksa
// "messages" alanına abone olunmamış. Veri tabanı varsa mesaj_senk'e yazılır,
// yoksa yalnızca süreç belleğinde tutulur (soğuk başlangıçta silinir).

import { dbConfigured, senkOku, senkYaz } from "./db";

export type IzKanal = "whatsapp" | "instagram";
export interface Iz { at: string; tur: string; ozet: string }

const bellek = new Map<string, Iz>();

export interface IzOzeti { son: Iz | null; kabul: Iz | null; red: Iz | null; islem: Iz | null }

/**
 * Üç kova: kabul edilen olay, reddedilen (imza) olay ve işleme sonucu ("islem": kaç yeni mesaj yazıldı,
 * "islem-hata": işlerken çıkan hata). Biri diğerini ezmez; kartta üçü de görünür.
 */
export async function izKaydet(kanal: IzKanal, tur: string, ozet: string): Promise<void> {
  const iz: Iz = { at: new Date().toISOString(), tur, ozet: ozet.slice(0, 400) };
  const kova = tur === "imza-red" ? "red" : tur.startsWith("islem") ? "islem" : "kabul";
  bellek.set(`${kanal}:${kova}`, iz);
  if (!dbConfigured()) return;
  try { await senkYaz(`iz:${kanal}:${kova}`, JSON.stringify(iz)); } catch { /* iz kaybı önemsiz */ }
}

async function tekOku(anahtar: string): Promise<Iz | null> {
  if (dbConfigured()) {
    try {
      const r = await senkOku(`iz:${anahtar}`);
      if (r?.deger) return JSON.parse(r.deger) as Iz;
    } catch { /* belleğe düş */ }
  }
  return bellek.get(anahtar) || null;
}

export async function izOku(kanal: IzKanal): Promise<IzOzeti> {
  const [kabul, red, islem] = await Promise.all([tekOku(`${kanal}:kabul`), tekOku(`${kanal}:red`), tekOku(`${kanal}:islem`)]);
  const son = kabul && red ? (new Date(kabul.at) >= new Date(red.at) ? kabul : red) : kabul || red;
  return { son, kabul, red, islem };
}
