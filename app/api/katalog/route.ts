// Toptan fiyat listesi — yalnız oturumu olanlara.
// Liste artık istemci JS paketinde değil: sipariş formu, fiyat listesi ve
// maliyet ekranı bu uçtan çeker. Oturumu olmayan (ör. paketi indiren
// herhangi biri) fiyatları göremez.
import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { FRAME_PROFILES } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET() {
  const user = await getSessionUser();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Giriş gerekli." }, { status: 401 });
  }
  return NextResponse.json({
    ok: true,
    profiles: FRAME_PROFILES,
    technical: TECHNICAL_PRODUCTS,
  });
}
