import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner, mesajUsernames } from "@/data/users";
import { dbConfigured, dbSaglik } from "@/lib/mesaj/db";
import { gmailConfigured, gmailHesaplar, gmailSenk, gmailTest } from "@/lib/mesaj/gmail";
import { instagramConfigured, instagramDurum } from "@/lib/mesaj/instagram";
import { serbestSablonAdlari, whatsappConfigured, whatsappDurum } from "@/lib/mesaj/whatsapp";
import { taslakHazir } from "@/lib/mesaj/taslak";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Mesajlar kurulum durumu ve canlı bağlantı testleri — yalnızca sahipler (Ayarlar sayfası).
async function yetki() {
  const u = await getSessionUser();
  return u && u.role === "staff" && isOwner(u.username) ? u : null;
}

export async function GET(req: NextRequest) {
  if (!(await yetki())) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  const canli = req.nextUrl.searchParams.get("test") === "1";
  const [db, gmail, wa, ig] = await Promise.all([
    canli ? dbSaglik() : Promise.resolve(null),
    canli && gmailConfigured() ? gmailTest() : Promise.resolve(null),
    canli && whatsappConfigured() ? whatsappDurum() : Promise.resolve(null),
    canli && instagramConfigured() ? instagramDurum() : Promise.resolve(null),
  ]);
  return NextResponse.json({
    ok: true,
    db: { kurulu: dbConfigured(), test: db },
    gmail: { kurulu: gmailConfigured(), hesaplar: gmailHesaplar().map((h) => h.adres), test: gmail },
    whatsapp: { kurulu: whatsappConfigured(), sablonlar: serbestSablonAdlari(), test: wa },
    instagram: { kurulu: instagramConfigured(), test: ig },
    taslak: taslakHazir(),
    gorebilenler: mesajUsernames(),
  });
}

// { islem: "gmail-senk" } → e-postaları şimdi çek
export async function POST(req: NextRequest) {
  if (!(await yetki())) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  const b = (await req.json().catch(() => null)) as { islem?: string } | null;
  if (b?.islem === "gmail-senk") {
    if (!dbConfigured()) return NextResponse.json({ ok: false, error: "Önce veri tabanı bağlanmalı (DATABASE_URL)." }, { status: 503 });
    if (!gmailConfigured()) return NextResponse.json({ ok: false, error: "GMAIL_HESAPLAR tanımlı değil." }, { status: 503 });
    const sonuc = await gmailSenk({ zorla: true });
    return NextResponse.json({ ok: true, sonuc });
  }
  return NextResponse.json({ ok: false, error: "Bilinmeyen işlem." }, { status: 400 });
}
