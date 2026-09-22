import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner, mesajUsernames } from "@/data/users";
import { dbConfigured, dbSaglik } from "@/lib/mesaj/db";
import { gmailConfigured, gmailHesaplar, gmailSenk, gmailTest } from "@/lib/mesaj/gmail";
import { igSayfaAbonelik, instagramConfigured, instagramDurum } from "@/lib/mesaj/instagram";
import { serbestSablonAdlari, wabaAbonelik, whatsappConfigured, whatsappDurum } from "@/lib/mesaj/whatsapp";
import { taslakHazir } from "@/lib/mesaj/taslak";
import { izOku } from "@/lib/mesaj/webhook-iz";

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
  const [db, gmail, wa, ig, waIz, igIz, abonelik] = await Promise.all([
    canli ? dbSaglik() : Promise.resolve(null),
    canli && gmailConfigured() ? gmailTest() : Promise.resolve(null),
    canli && whatsappConfigured() ? whatsappDurum() : Promise.resolve(null),
    canli && instagramConfigured() ? instagramDurum() : Promise.resolve(null),
    izOku("whatsapp"),
    izOku("instagram"),
    canli && whatsappConfigured() && process.env.WHATSAPP_WABA_ID ? wabaAbonelik(false) : Promise.resolve(null),
  ]);
  const igAbonelik = canli && instagramConfigured() && (process.env.INSTAGRAM_PAGE_ID || "").trim() ? await igSayfaAbonelik(false) : null;
  const origin = req.nextUrl.origin;
  // Yayındaki secret'ın kısa parmak izi (tamamı asla dönmez): Meta'daki değerle karşılaştırmak için.
  const secret = (process.env.WHATSAPP_APP_SECRET || process.env.META_APP_SECRET || "").trim();
  const secretIpucu = secret ? { uzunluk: secret.length, bas: secret.slice(0, 2), son: secret.slice(-2), tirnak: /^["']|["']$/.test(secret) } : null;
  return NextResponse.json({
    ok: true,
    db: { kurulu: dbConfigured(), test: db },
    gmail: { kurulu: gmailConfigured(), hesaplar: gmailHesaplar().map((h) => h.adres), test: gmail },
    whatsapp: {
      kurulu: whatsappConfigured(), sablonlar: serbestSablonAdlari(), test: wa,
      webhook: { url: `${origin}/api/whatsapp/webhook`, verifyToken: Boolean((process.env.WHATSAPP_VERIFY_TOKEN || "").trim()), appSecret: Boolean(secret), secretIpucu, sonOlay: waIz.son, kabul: waIz.kabul, red: waIz.red, wabaId: Boolean((process.env.WHATSAPP_WABA_ID || "").trim()), abonelik },
    },
    instagram: { kurulu: instagramConfigured(), test: ig, webhook: { url: `${origin}/api/mesaj/webhook`, sonOlay: igIz.son, kabul: igIz.kabul, red: igIz.red, abonelik: igAbonelik, sayfa: Boolean((process.env.INSTAGRAM_PAGE_ID || "").trim()) } },
    taslak: taslakHazir(),
    gorebilenler: mesajUsernames(),
  });
}

// { islem: "gmail-senk" } → e-postaları şimdi çek
export async function POST(req: NextRequest) {
  if (!(await yetki())) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  const b = (await req.json().catch(() => null)) as { islem?: string } | null;
  if (b?.islem === "ig-abone") {
    const r = await igSayfaAbonelik(true);
    return NextResponse.json({ ok: r.ok, abonelik: r, error: r.ok ? undefined : r.hata });
  }
  if (b?.islem === "wa-abone") {
    const r = await wabaAbonelik(true);
    return NextResponse.json({ ok: r.ok, abonelik: r, error: r.ok ? undefined : r.hata });
  }
  if (b?.islem === "gmail-senk") {
    if (!dbConfigured()) return NextResponse.json({ ok: false, error: "Önce veri tabanı bağlanmalı (DATABASE_URL)." }, { status: 503 });
    if (!gmailConfigured()) return NextResponse.json({ ok: false, error: "GMAIL_HESAPLAR tanımlı değil." }, { status: 503 });
    const sonuc = await gmailSenk({ zorla: true });
    return NextResponse.json({ ok: true, sonuc });
  }
  return NextResponse.json({ ok: false, error: "Bilinmeyen işlem." }, { status: 400 });
}
