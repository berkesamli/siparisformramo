// Duyuru süzme — saf fonksiyon, sunucudan da istemciden de çağrılabilir.

import { DUYURULAR, type Duyuru } from "@/data/duyurular";

const SONSUZ = "9999-12-31";

/** Verilen rol ve gün (YYYY-MM-DD, İstanbul) için aktif duyurular; en yeni önce. */
export function aktifDuyurular(role: "staff" | "customer", today: string): Duyuru[] {
  return DUYURULAR
    .filter((d) => d.roller.includes(role) && d.baslangic <= today && today <= (d.bitis ?? SONSUZ))
    .sort((a, b) => (a.baslangic < b.baslangic ? 1 : a.baslangic > b.baslangic ? -1 : 0));
}
