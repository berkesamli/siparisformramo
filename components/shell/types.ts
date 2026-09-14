import type { NavFlags } from "./nav-config";

export interface ShellUser extends NavFlags {
  name: string;
  username: string;
}

/** /api/dashboard?lite=1 — kenar çubuğu ve rozetler için hafif sayaçlar. */
export interface ShellStats {
  bugun: number;          // bugün girilen sipariş (toptan + perakende)
  bugunToptan: number;
  bugunPerakende: number;
  dun: number;            // dün girilen sipariş (delta için)
  acik: number;           // açık toptan sipariş (oluşturuldu + hazırlanıyor)
  perakendeAcik: number;  // açık perakende sipariş
  kontrolsuz: number;     // son 7 günde kontrol edilmemiş toptan sipariş
  ayAdet: number;         // bu ay toplam sipariş adedi (toptan + perakende)
  ayCiro?: number;        // bu ay toptan ciro (yalnızca finans yetkisi)
  blob: boolean;          // depo bağlı mı (değilse sayaçlar 0 gelir)
}

export const STATS_KEY = "olga-shell-stats";
export const STATS_TTL_MS = 3 * 60_000;
