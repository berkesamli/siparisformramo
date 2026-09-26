import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { notGuncelle, notSil } from "@/lib/uretim/db";
import { s, subeOku, tarihOku } from "@/lib/uretim/girdi";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function PATCH(req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as any;
  if (!b) return NextResponse.json({ ok: false, error: "Geçersiz istek." }, { status: 400 });
  const d: Parameters<typeof notGuncelle>[1] = {};
  if (b.tarih !== undefined) { const t = tarihOku(b.tarih); if (!t) return NextResponse.json({ ok: false, error: "Geçersiz tarih." }, { status: 400 }); d.tarih = t; }
  if (b.sube !== undefined) d.sube = subeOku(b.sube) || "";
  if (b.metin !== undefined) { d.metin = s(b.metin, 1000).trim(); if (!d.metin) return NextResponse.json({ ok: false, error: "Not boş olamaz." }, { status: 400 }); }
  if (b.tamam !== undefined) d.tamam = Boolean(b.tamam);
  try {
    const not = await notGuncelle(params.id, d);
    if (!not) return NextResponse.json({ ok: false, error: "Not bulunamadı." }, { status: 404 });
    return NextResponse.json({ ok: true, not });
  } catch (e) {
    return hataYaniti(e);
  }
}

export async function DELETE(_req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  try {
    const ok = await notSil(params.id);
    return ok ? NextResponse.json({ ok: true }) : NextResponse.json({ ok: false, error: "Not bulunamadı." }, { status: 404 });
  } catch (e) {
    return hataYaniti(e);
  }
}
