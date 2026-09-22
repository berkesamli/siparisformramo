// Tek kapı: konuşmanın kanalına göre yanıtı gönderir ve mesaj olarak kaydeder.

import { konusma, konusmaGuncelle, mesajEkle } from "./db";
import { gmailConfigured, gmailGonder } from "./gmail";
import { instagramConfigured, instagramGonder } from "./instagram";
import { whatsappConfigured, whatsappGonder } from "./whatsapp";
import type { Kanal, Mesaj } from "./tur";

export function kanalDurumu(): Record<Kanal, boolean> {
  return { whatsapp: whatsappConfigured(), instagram: instagramConfigured(), email: gmailConfigured() };
}

export async function yanitGonder(
  konusmaId: string,
  metin: string,
  kullanici: { username: string; name: string },
  taslakAi = false
): Promise<Mesaj> {
  const k = await konusma(konusmaId);
  if (!k) throw new Error("Konuşma bulunamadı.");
  const govde = String(metin || "").trim();
  if (!govde) throw new Error("Boş mesaj gönderilemez.");
  if (govde.length > 4000 && k.kanal !== "email") throw new Error("Mesaj çok uzun (en fazla 4000 karakter).");

  let disId: string | null = null;
  if (k.kanal === "whatsapp") {
    disId = (await whatsappGonder(k, govde)).disId;
  } else if (k.kanal === "instagram") {
    disId = (await instagramGonder(k, govde)).disId;
  } else {
    const r = await gmailGonder(k, govde);
    disId = r.disId;
    await konusmaGuncelle(k.id, { meta: { ...k.meta, references: r.references } });
  }
  const { mesaj } = await mesajEkle({ konusmaId: k.id, yon: "giden", govde, disId, gonderen: kullanici.name, taslakAi, durum: "gonderildi" });
  return mesaj;
}
