import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isFinance, isKurYetkili, isOwner } from "@/data/users";
import { computeStaffDashboard, computeCustomerDashboard } from "@/lib/dashboard";
import { memo } from "@/lib/server-cache";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

// Gösterge paneli verisi. ?lite=1 → yalnızca kenar çubuğu sayaçları.
// Sonuç yetki kümesine göre 45 sn süreç içi önbellekte tutulur.
export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Giriş gerekli." }, { status: 401 });
  }
  try {
    if (user.role !== "staff") {
      const data = await memo("dash:customer", 45_000, () => computeCustomerDashboard());
      return NextResponse.json({ ok: true, data });
    }
    const flags = {
      finance: isFinance(user.username),
      kur: isKurYetkili(user.username),
      owner: isOwner(user.username),
    };
    const key = `dash:staff:${flags.finance ? 1 : 0}${flags.kur ? 1 : 0}${flags.owner ? 1 : 0}`;
    const data = await memo(key, 45_000, () => computeStaffDashboard(flags));
    if (req.nextUrl.searchParams.get("lite") === "1") {
      return NextResponse.json({ ok: true, lite: data.lite });
    }
    return NextResponse.json({ ok: true, data });
  } catch (err) {
    console.error("Gösterge paneli hesaplanamadı:", err);
    return NextResponse.json({ ok: false, error: "Veriler alınamadı." }, { status: 500 });
  }
}
