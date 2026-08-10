import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { getDailyRates, saveDailyRates, istanbulDateKey } from "@/lib/orders";
import { isOwner } from "@/data/users";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Bugünün kuru — sipariş formu açılınca otomatik doldurmak için.
// Yanıt, isteği yapanın kuru değiştirme yetkisi olup olmadığını da söyler.
export async function GET() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const rates = await getDailyRates(istanbulDateKey());
  return NextResponse.json({
    ok: true,
    rates,
    dateKey: istanbulDateKey(),
    yetkili: isOwner(user.username),
  });
}

// Günün kurunu belirleme — yalnızca firma sahipleri (OWNER_USERNAMES).
// Gün içinde kur bir kez girildikten sonra diğer çalışanların sipariş
// formunda kur alanı kilitlenir; herkes aynı kurdan sipariş girer.
export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  if (!isOwner(user.username)) {
    return NextResponse.json(
      { ok: false, error: "Günlük kuru yalnızca yetkili kişiler girebilir." },
      { status: 403 }
    );
  }
  const body = await req.json().catch(() => null);
  const rate = Number(body?.rate) || 0;
  const euroRate = Number(body?.euroRate) || 0;
  if (rate <= 0 && euroRate <= 0) {
    return NextResponse.json({ ok: false, error: "Kur değeri gerekli." }, { status: 400 });
  }
  const rates = {
    rate,
    euroRate,
    updatedAt: new Date().toISOString(),
    by: user.name,
    sabit: true,
  };
  await saveDailyRates(istanbulDateKey(), rates);
  return NextResponse.json({ ok: true, rates });
}
