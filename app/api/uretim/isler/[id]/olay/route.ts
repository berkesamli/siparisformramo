import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { isBul, olayEkle } from "@/lib/uretim/db";
import { s } from "@/lib/uretim/girdi";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// İşe not (olay) ekler — zaman çizelgesinde görünür.
export async function POST(req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as { aciklama?: string } | null;
  const aciklama = s(b?.aciklama, 2000).trim();
  if (!aciklama) return NextResponse.json({ ok: false, error: "Not boş olamaz." }, { status: 400 });
  try {
    const is = await isBul(params.id);
    if (!is) return NextResponse.json({ ok: false, error: "İş bulunamadı." }, { status: 404 });
    const olay = await olayEkle(is.id, "not", aciklama, y.user.name);
    return NextResponse.json({ ok: true, olay });
  } catch (e) {
    return hataYaniti(e);
  }
}
