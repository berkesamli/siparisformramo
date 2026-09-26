import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { isBul, isGuncelle, isSil, olaylar, type IsDegisiklik } from "@/lib/uretim/db";
import { kalemleriOku, r2, s, saatOku, subeOku, tarihOku, durumOku } from "@/lib/uretim/girdi";
import { perakendeDurumYaz } from "@/lib/uretim/perakende-senk";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

export async function GET(_req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  try {
    const is = await isBul(params.id);
    if (!is) return NextResponse.json({ ok: false, error: "İş bulunamadı." }, { status: 404 });
    return NextResponse.json({ ok: true, is, olaylar: await olaylar(params.id) });
  } catch (e) {
    return hataYaniti(e);
  }
}

// PATCH: durum / şube / plan tarihi-saati / teslim tarihi / notlar / müşteri / kalemler.
// Perakende kaynaklı işte durum değişince sipariş kaydı da güncellenir.
export async function PATCH(req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as any;
  if (!b) return NextResponse.json({ ok: false, error: "Geçersiz istek." }, { status: 400 });
  const d: IsDegisiklik = {};
  if (b.durum !== undefined) { const v = durumOku(b.durum); if (!v) return NextResponse.json({ ok: false, error: "Geçersiz durum." }, { status: 400 }); d.durum = v; }
  if (b.sube !== undefined) d.sube = subeOku(b.sube);
  if (b.planTarih !== undefined) { const v = tarihOku(b.planTarih); if (v === undefined) return NextResponse.json({ ok: false, error: "Geçersiz plan tarihi." }, { status: 400 }); d.planTarih = v; }
  if (b.teslimTarih !== undefined) { const v = tarihOku(b.teslimTarih); if (v === undefined) return NextResponse.json({ ok: false, error: "Geçersiz teslim tarihi." }, { status: 400 }); d.teslimTarih = v; }
  if (b.planSaat !== undefined) { const v = saatOku(b.planSaat); if (v === undefined) return NextResponse.json({ ok: false, error: "Geçersiz saat." }, { status: 400 }); d.planSaat = v; }
  if (b.planSira !== undefined) d.planSira = Number(b.planSira) || 0;
  if (b.notlar !== undefined) d.notlar = s(b.notlar, 2000).trim();
  if (b.baslik !== undefined) d.baslik = s(b.baslik, 160).trim();
  if (b.musteriAd !== undefined) d.musteriAd = s(b.musteriAd, 120).trim();
  if (b.musteriTel !== undefined) d.musteriTel = s(b.musteriTel, 40).trim();
  if (b.musteriAdres !== undefined) d.musteriAdres = s(b.musteriAdres, 240).trim();
  if (b.musteriSehir !== undefined) d.musteriSehir = s(b.musteriSehir, 60).trim();
  if (b.kalemler !== undefined) d.kalemler = kalemleriOku(b.kalemler);
  if (b.tutar !== undefined) d.tutar = r2(b.tutar);
  if (b.subeOneriUygula === true) {
    const eski = await isBul(params.id);
    if (eski?.subeOneri) d.sube = eski.subeOneri;
  }
  try {
    const is = await isGuncelle(params.id, d, y.user.name);
    if (!is) return NextResponse.json({ ok: false, error: "İş bulunamadı." }, { status: 404 });
    let kaynakGuncellendi = false;
    if (d.durum && is.kaynak === "perakende") {
      try { kaynakGuncellendi = await perakendeDurumYaz(is.id, d.durum); } catch (e) { console.error("Perakende durum yazılamadı:", e); }
    }
    return NextResponse.json({ ok: true, is, kaynakGuncellendi });
  } catch (e) {
    console.error("İş güncellenemedi:", e);
    return hataYaniti(e);
  }
}

// DELETE: yalnızca elle/teklif kaynaklı işler silinir; kaynağı olan işler iptal edilir.
export async function DELETE(_req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  try {
    const is = await isBul(params.id);
    if (!is) return NextResponse.json({ ok: false, error: "İş bulunamadı." }, { status: 404 });
    if (is.kaynak === "perakende" || is.kaynak === "online") {
      await isGuncelle(is.id, { durum: "iptal" }, y.user.name);
      return NextResponse.json({ ok: true, iptal: true });
    }
    await isSil(is.id);
    return NextResponse.json({ ok: true, silindi: true });
  } catch (e) {
    return hataYaniti(e);
  }
}
