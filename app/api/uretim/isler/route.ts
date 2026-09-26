import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { isEkle, isler } from "@/lib/uretim/db";
import { kalemleriOku, r2, s, saatOku, subeOku, tarihOku, durumOku } from "@/lib/uretim/girdi";
import { TARIH_RE, type IsDurum, type IsKaynak, type Sube } from "@/lib/uretim/tur";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

// GET: liste (arama / arşiv). Takvim için /api/uretim/takvim kullanılır.
export async function GET(req: NextRequest) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const q = req.nextUrl.searchParams;
  const from = q.get("from") || undefined, to = q.get("to") || undefined;
  if ((from && !TARIH_RE.test(from)) || (to && !TARIH_RE.test(to))) {
    return NextResponse.json({ ok: false, error: "Geçersiz tarih." }, { status: 400 });
  }
  try {
    const liste = await isler({
      from, to,
      plansiz: q.get("plansiz") === "1",
      kaynak: (q.get("kaynak") || "") as IsKaynak | "",
      sube: (q.get("sube") || "hepsi") as Sube | "hepsi",
      durum: (q.get("durum") || "") as IsDurum | "acik" | "kapali" | "",
      q: q.get("q") || undefined,
      limit: Number(q.get("limit")) || undefined,
    });
    return NextResponse.json({ ok: true, isler: liste });
  } catch (e) {
    return hataYaniti(e);
  }
}

// POST: elle iş ekle (mağazaya gelen, telefonla alınan vb.)
export async function POST(req: NextRequest) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as any;
  if (!b) return NextResponse.json({ ok: false, error: "Geçersiz istek." }, { status: 400 });
  const musteriAd = s(b.musteriAd, 120).trim();
  const baslik = s(b.baslik, 160).trim() || musteriAd;
  if (!baslik) return NextResponse.json({ ok: false, error: "Başlık ya da müşteri adı gerekli." }, { status: 400 });
  const planTarih = tarihOku(b.planTarih);
  const teslimTarih = tarihOku(b.teslimTarih);
  if (planTarih === undefined && b.planTarih !== undefined) return NextResponse.json({ ok: false, error: "Geçersiz plan tarihi." }, { status: 400 });
  if (teslimTarih === undefined && b.teslimTarih !== undefined) return NextResponse.json({ ok: false, error: "Geçersiz teslim tarihi." }, { status: 400 });
  const planSaat = saatOku(b.planSaat);
  if (planSaat === undefined && b.planSaat !== undefined) return NextResponse.json({ ok: false, error: "Geçersiz saat." }, { status: 400 });
  const kaynak: IsKaynak = ["elle", "toptan"].includes(String(b.kaynak)) ? b.kaynak : "elle";
  try {
    const is = await isEkle({
      kaynak,
      kaynakRef: s(b.kaynakRef, 60).trim(),
      baslik, musteriAd,
      musteriTel: s(b.musteriTel, 40).trim(),
      musteriAdres: s(b.musteriAdres, 240).trim(),
      musteriSehir: s(b.musteriSehir, 60).trim(),
      sube: subeOku(b.sube) || "",
      planTarih: planTarih ?? null,
      planSaat: planSaat || "",
      teslimTarih: teslimTarih ?? null,
      durum: durumOku(b.durum),
      kalemler: kalemleriOku(b.kalemler),
      tutar: r2(b.tutar),
      notlar: s(b.notlar, 2000).trim(),
    }, y.user.name);
    return NextResponse.json({ ok: true, is });
  } catch (e) {
    console.error("İş eklenemedi:", e);
    return hataYaniti(e);
  }
}
