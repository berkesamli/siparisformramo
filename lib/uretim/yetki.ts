// Üretim takvimi rotaları için ortak yetki denetimi: oturum + çalışan +
// üretim yetkisi (sahipler + URETIM_USERNAMES) + veri tabanı hazır mı.
import { NextResponse } from "next/server";
import { getSessionUser, type SessionUser } from "@/lib/auth";
import { isUretimci } from "@/data/users";
import { dbConfigured } from "./db";

export async function uretimKullanici(): Promise<{ user: SessionUser } | { hata: NextResponse }> {
  const user = await getSessionUser();
  if (!user) return { hata: NextResponse.json({ ok: false, error: "Giriş gerekli." }, { status: 401 }) };
  if (user.role !== "staff" || !isUretimci(user.username)) {
    return { hata: NextResponse.json({ ok: false, error: "Üretim takvimi için yetkiniz yok." }, { status: 403 }) };
  }
  if (!dbConfigured()) {
    return { hata: NextResponse.json({ ok: false, kurulum: true, error: "Üretim takvimi için veri tabanı bağlı değil (DATABASE_URL)." }, { status: 503 }) };
  }
  return { user };
}

/** Cron (Authorization: Bearer CRON_SECRET) ya da yetkili çalışan. */
export async function cronYaDaUretimci(authHeader: string | null): Promise<{ ok: true; by: string } | { ok: false; hata: NextResponse }> {
  const cron = (process.env.CRON_SECRET || "").trim();
  if (cron && (authHeader || "") === `Bearer ${cron}`) return { ok: true, by: "cron" };
  const u = await getSessionUser();
  if (u && u.role === "staff" && isUretimci(u.username)) return { ok: true, by: u.name };
  return { ok: false, hata: NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 }) };
}

export function hataYaniti(e: unknown, durum = 500): NextResponse {
  const msg = (e as Error)?.message || String(e);
  return NextResponse.json({ ok: false, error: msg }, { status: durum });
}
