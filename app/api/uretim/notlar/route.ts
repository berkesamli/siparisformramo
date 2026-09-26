import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { notEkle, notlar } from "@/lib/uretim/db";
import { s, subeOku, tarihOku } from "@/lib/uretim/girdi";
import { TARIH_RE } from "@/lib/uretim/tur";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(req: NextRequest) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const from = req.nextUrl.searchParams.get("from") || "", to = req.nextUrl.searchParams.get("to") || "";
  if (!TARIH_RE.test(from) || !TARIH_RE.test(to)) return NextResponse.json({ ok: false, error: "Geçersiz aralık." }, { status: 400 });
  try {
    return NextResponse.json({ ok: true, notlar: await notlar(from, to) });
  } catch (e) {
    return hataYaniti(e);
  }
}

// Gün notu ekle: { tarih, sube?, metin }
export async function POST(req: NextRequest) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as any;
  const tarih = tarihOku(b?.tarih);
  const metin = s(b?.metin, 1000).trim();
  if (!tarih || !metin) return NextResponse.json({ ok: false, error: "Tarih ve not metni gerekli." }, { status: 400 });
  try {
    const not = await notEkle({ tarih, sube: subeOku(b?.sube) || "", metin }, y.user.name);
    return NextResponse.json({ ok: true, not });
  } catch (e) {
    return hataYaniti(e);
  }
}
