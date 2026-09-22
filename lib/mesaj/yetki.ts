// Mesaj rotaları için ortak yetki denetimi: oturum + çalışan + mesaj yetkisi.
import { NextResponse } from "next/server";
import { getSessionUser, type SessionUser } from "@/lib/auth";
import { isMesajci } from "@/data/users";
import { dbConfigured } from "./db";

export async function mesajKullanici(): Promise<{ user: SessionUser } | { hata: NextResponse }> {
  const user = await getSessionUser();
  if (!user) return { hata: NextResponse.json({ ok: false, error: "Giriş gerekli." }, { status: 401 }) };
  if (user.role !== "staff" || !isMesajci(user.username)) {
    return { hata: NextResponse.json({ ok: false, error: "Bu ekran için yetkiniz yok." }, { status: 403 }) };
  }
  if (!dbConfigured()) {
    return { hata: NextResponse.json({ ok: false, kurulum: true, error: "Mesajlar için veri tabanı bağlı değil (DATABASE_URL)." }, { status: 503 }) };
  }
  return { user };
}

export function hataYaniti(e: unknown, durum = 500): NextResponse {
  const msg = (e as Error)?.message || String(e);
  return NextResponse.json({ ok: false, error: msg }, { status: durum });
}
