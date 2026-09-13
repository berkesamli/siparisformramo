// Giriş kilidi (brute-force önlemi): aynı kullanıcı adına 15 dakika içinde
// 5 hatalı parola girilirse hesap 15 dakika kilitlenir. Sayaç Vercel Blob'da
// tutulur (sunucusuz ortamda bellek her istekte sıfırlanabilir); Blob yoksa
// süreç içi bellekle idare edilir (yerel geliştirme).

import { blobConfigured } from "./orders";
import { normalizeUsername } from "@/data/users";

const MAX_FAILS = 5;
const WINDOW_MS = 15 * 60 * 1000;
const LOCK_MS = 15 * 60 * 1000;

interface FailRecord {
  fails: number;
  firstAt: number; // epoch ms — pencerenin başı
  lockedUntil: number; // epoch ms — 0 = kilitli değil
}

const bellek = new Map<string, FailRecord>();

const kayitYolu = (key: string) => `security/login/${key}.json`;

async function readRecord(key: string): Promise<FailRecord | null> {
  if (!blobConfigured()) return bellek.get(key) || null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(kayitYolu(key), { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as FailRecord;
  } catch {
    return bellek.get(key) || null;
  }
}

async function writeRecord(key: string, rec: FailRecord): Promise<void> {
  bellek.set(key, rec);
  if (!blobConfigured()) return;
  try {
    const { put } = await import("@vercel/blob");
    await put(kayitYolu(key), JSON.stringify(rec), {
      access: "private",
      contentType: "application/json",
      addRandomSuffix: false,
      allowOverwrite: true,
    });
  } catch {
    /* Blob yazılamazsa bellek kaydı yeter — girişi engelleme */
  }
}

/** Kilitliyse kalan saniyeyi döner; değilse 0. */
export async function loginLockRemaining(username: string): Promise<number> {
  const key = normalizeUsername(username);
  if (!key) return 0;
  const rec = await readRecord(key);
  if (!rec || !rec.lockedUntil) return 0;
  const kalan = rec.lockedUntil - Date.now();
  return kalan > 0 ? Math.ceil(kalan / 1000) : 0;
}

/** Hatalı girişi kaydeder; hesap yeni kilitlendiyse true döner. */
export async function registerLoginFail(username: string): Promise<boolean> {
  const key = normalizeUsername(username);
  if (!key) return false;
  const now = Date.now();
  const eski = await readRecord(key);
  let rec: FailRecord;
  if (!eski || now - eski.firstAt > WINDOW_MS) {
    rec = { fails: 1, firstAt: now, lockedUntil: 0 };
  } else {
    rec = { ...eski, fails: eski.fails + 1 };
  }
  const yeniKilit = rec.fails >= MAX_FAILS && !(eski && eski.lockedUntil > now);
  if (rec.fails >= MAX_FAILS) {
    rec.lockedUntil = now + LOCK_MS;
    rec.fails = 0;
    rec.firstAt = now;
  }
  await writeRecord(key, rec);
  return yeniKilit;
}

/** Başarılı girişte sayaç sıfırlanır. */
export async function clearLoginFails(username: string): Promise<void> {
  const key = normalizeUsername(username);
  if (!key) return;
  bellek.delete(key);
  if (!blobConfigured()) return;
  try {
    const { del } = await import("@vercel/blob");
    await del(kayitYolu(key));
  } catch {
    /* silinemezse pencere süresi dolunca kendiliğinden geçersizleşir */
  }
}
