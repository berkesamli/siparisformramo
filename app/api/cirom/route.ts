import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { computePersonalSales } from "@/lib/personal-sales";
import { memo } from "@/lib/server-cache";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

// Çalışanın kendi satışları. ?ay=YYYY-MM ile son 6 ay içinden ay seçilir.
// Sonuç istek sahibinin adıyla girilmiş siparişleri içerir; bölge sorumlusuysa
// bölgesindeki müşterilerin tüm satışları da eklenir (ad oturumdan gelir).
export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const ay = req.nextUrl.searchParams.get("ay") || "";
  const ayKey = /^\d{4}-\d{2}$/.test(ay) ? ay : "";
  try {
    const data = await memo(`cirom:${user.username}:${ayKey || "bu-ay"}`, 30_000, () =>
      computePersonalSales(user.name, ayKey || undefined, user.username)
    );
    return NextResponse.json({ ok: true, data });
  } catch (err) {
    console.error("Kişisel satış özeti hesaplanamadı:", err);
    return NextResponse.json({ ok: false, error: "Veriler alınamadı." }, { status: 500 });
  }
}
