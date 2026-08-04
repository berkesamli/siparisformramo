// Çek/senet kayıtları — finans/ceksenet/<id>.json (düz liste; portföy küçük).
//
// Kasa kuralı: alınan çek/senet cariyi ALINDIĞINDA düşürür (Tahsilat kaydı,
// method cek/senet), kasaya ise ancak TAHSİL edildiğinde girer (banka).
// Ciro kasaya hiç dokunmaz — bir tedarikçi borcunu kapatır, çek el değiştirir.
//
// "Çek kimde?" cevabı durumdan okunur:
//   portfoyde → bizde · ciro → ciroTarget'ta · tahsil → bankada tahsil edildi
//   karsiliksiz → sorunlu takip · iade → müşteriye geri verildi

import { blobConfigured } from "./orders";
import type { Branch } from "./customers";
import type { FinansKaynak } from "./tahsilat";

export type CekSenetTur = "alinan" | "verilen";
export type CekSenetKind = "cek" | "senet";
export type CekSenetDurum =
  | "portfoyde"
  | "tahsil"
  | "ciro"
  | "karsiliksiz"
  | "odendi"
  | "iade";

export const CEKSENET_DURUM_LABELS: Record<CekSenetDurum, string> = {
  portfoyde: "Portföyde",
  tahsil: "Tahsil Edildi",
  ciro: "Ciro Edildi",
  karsiliksiz: "Karşılıksız",
  odendi: "Ödendi",
  iade: "İade",
};

// Geçerli durum geçişleri — alınan ve verilen için ayrı akış.
export function allowedTransitions(
  tur: CekSenetTur,
  durum: CekSenetDurum
): CekSenetDurum[] {
  if (durum !== "portfoyde") return []; // terminal durumlar
  return tur === "alinan"
    ? ["tahsil", "ciro", "karsiliksiz", "iade"]
    : ["odendi", "iade"];
}

export interface CekSenetDurumEvent {
  durum: CekSenetDurum;
  date: string; // YYYY-MM-DD işlem tarihi
  by: string;
  note?: string;
}

export interface CekSenet {
  id: string; // CS-YYYYMMDDHHMMSS-xxxxx
  createdAt: string;
  createdBy: string;
  tur: CekSenetTur;
  kind: CekSenetKind;
  branch: Branch;
  banka?: string; // senette boş kalabilir
  bankaSube?: string;
  hesapNo?: string;
  belgeNo?: string; // çek/senet numarası
  /** Keşideci — çeki fiilen yazan kişi/firma (üçüncü şahıs çeklerinde müşteriden farklı). */
  cekSahibi?: string;
  tutar: number; // TL
  vade: string; // YYYY-MM-DD
  customerId?: string; // alınan: kimden
  customerName?: string;
  supplier?: string; // verilen: kime
  durum: CekSenetDurum;
  ciroTarget?: string; // ciro edildiyse hangi tedarikçiye
  ciroDate?: string;
  tahsilDate?: string; // bankadan tahsil tarihi — kasa girişi bu tarihte
  history: CekSenetDurumEvent[];
  tahsilatId?: string; // alınan çekin girişinde oluşan Tahsilat kaydı
  giderId?: string; // verilen çek bir gidere bağlıysa
  note?: string;
  kaynak: FinansKaynak;
}

const path = (id: string) => `finans/ceksenet/${id}.json`;

export function newCekSenetId(now = new Date()): string {
  const t = now.toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  return `CS-${t}-${Math.random().toString(36).slice(2, 7)}`;
}

export async function saveCekSenet(c: CekSenet): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(path(c.id), JSON.stringify(c), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function getCekSenet(id: string): Promise<CekSenet | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(path(id), { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as CekSenet;
  } catch {
    return null;
  }
}

export async function listCekSenet(limit = 2000): Promise<CekSenet[]> {
  if (!blobConfigured()) return [];
  const { list, get } = await import("@vercel/blob");
  const out: CekSenet[] = [];
  try {
    const { blobs } = await list({ prefix: "finans/ceksenet/", limit });
    await Promise.all(
      blobs.map(async (b) => {
        try {
          const r = await get(b.pathname, {
            access: "private",
            useCache: false,
          });
          if (!r || r.statusCode !== 200 || !r.stream) return;
          out.push(JSON.parse(await new Response(r.stream).text()) as CekSenet);
        } catch {
          /* tek kayıt okunamazsa listeyi bozma */
        }
      })
    );
  } catch {
    /* boş */
  }
  // Vadesi yakın olan üstte
  return out.sort((a, b) => a.vade.localeCompare(b.vade));
}
