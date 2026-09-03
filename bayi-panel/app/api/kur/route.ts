import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { getUsdRate } from "@/lib/kur";

export const dynamic = "force-dynamic";

// Günün USD kuru (TCMB). Bulunamazsa rate: null — bayi elle girer.
export async function GET() {
  const user = await getSessionUser();
  if (!user) return NextResponse.json({ ok: false, error: "Yetkisiz" }, { status: 401 });
  const r = await getUsdRate();
  return NextResponse.json({ ok: true, rate: r?.usd ?? null, dateKey: r?.dateKey ?? null });
}
