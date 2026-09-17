// Müşteriye WhatsApp ile gönderilen fişlerin teslim takibi.
//
// Meta, mesajı KABUL ettiğinde bir wamid döner; numarada WhatsApp yoksa
// teslim hatası ("failed", kod 131026) ancak webhook'la sonradan gelir.
// Bu yüzden gönderim anında wamid → sipariş/telefon/SMS metni kaydı
// yazılır; webhook "failed" görünce kayda bakar ve SMS'e düşer,
// "delivered/read" görünce kaydı siler.
//
// Depo: wa/bekleyen/<sha256(wamid)>.json

import { createHash } from "node:crypto";
import { blobConfigured } from "./orders";

export interface BekleyenTeslim {
  wamid: string;
  orderId: string;
  dateKey: string;
  tur: "toptan" | "perakende";
  telefon: string;    // normalize (90…)
  musteri: string;
  smsMetni: string;   // WhatsApp başarısız olursa gidecek SMS
  createdAt: string;  // ISO
}

const anahtar = (wamid: string) => createHash("sha256").update(wamid).digest("hex").slice(0, 40);
const yol = (wamid: string) => `wa/bekleyen/${anahtar(wamid)}.json`;

export async function bekleyenKaydet(rec: BekleyenTeslim): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(yol(rec.wamid), JSON.stringify(rec), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function bekleyenAl(wamid: string): Promise<BekleyenTeslim | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(yol(wamid), { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    const text = await new Response(r.stream).text();
    return JSON.parse(text) as BekleyenTeslim;
  } catch {
    return null;
  }
}

export async function bekleyenSil(wamid: string): Promise<void> {
  if (!blobConfigured()) return;
  try {
    const { del } = await import("@vercel/blob");
    await del(yol(wamid));
  } catch {
    /* yoksa sorun değil */
  }
}
