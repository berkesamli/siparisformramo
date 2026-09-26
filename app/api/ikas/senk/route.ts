import { NextRequest, NextResponse } from "next/server";
import { cronYaDaUretimci, hataYaniti } from "@/lib/uretim/yetki";
import { dbConfigured } from "@/lib/uretim/db";
import { ikasSenk } from "@/lib/ikas/senk";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 120;

// ikas senkronu: cron (Authorization: Bearer CRON_SECRET) ya da yetkili çalışan. ?tam=1 → son 90 gün.
export async function GET(req: NextRequest) {
  const y = await cronYaDaUretimci(req.headers.get("authorization"));
  if (!y.ok) return y.hata;
  if (!dbConfigured()) return NextResponse.json({ ok: false, error: "DATABASE_URL yok." }, { status: 503 });
  try {
    const sonuc = await ikasSenk(y.by, { tam: req.nextUrl.searchParams.get("tam") === "1" });
    return NextResponse.json({ ok: !sonuc.hata, sonuc });
  } catch (e) {
    return hataYaniti(e);
  }
}
