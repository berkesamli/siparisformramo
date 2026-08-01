import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { getDailyRates, saveDailyRates, istanbulDateKey } from "@/lib/orders";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Bugünün kuru — sipariş formu açılınca otomatik doldurmak için
export async function GET() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const rates = await getDailyRates(istanbulDateKey());
  return NextResponse.json({ ok: true, rates });
}

// Günün kurunu elle güncelleme (formda kur değiştirilirse)
export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = await req.json().catch(() => null);
  const rate = Number(body?.rate) || 0;
  const euroRate = Number(body?.euroRate) || 0;
  if (rate <= 0 && euroRate <= 0) {
    return NextResponse.json({ ok: false, error: "Kur değeri gerekli." }, { status: 400 });
  }
  await saveDailyRates(istanbulDateKey(), {
    rate,
    euroRate,
    updatedAt: new Date().toISOString(),
    by: user.name,
  });
  return NextResponse.json({ ok: true });
}
