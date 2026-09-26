import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { isBul, isGuncelle, olayEkle } from "@/lib/uretim/db";
import { ekKaydet, ekSil, ekTuru, EK_AZAMI_BAYT } from "@/lib/uretim/ek";
import { blobConfigured } from "@/lib/orders";
import type { IsEk } from "@/lib/uretim/tur";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// POST (multipart, alan adı "dosya"): işe ek yükler. "foy=1" ile yüklenen PDF işin föyü olur
// (Claude'da üretilmiş föyü takvime iliştirmek için).
export async function POST(req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  if (!blobConfigured()) return NextResponse.json({ ok: false, error: "Dosya deposu (Blob) bağlı değil." }, { status: 503 });
  let form: FormData;
  try { form = await req.formData(); } catch { return NextResponse.json({ ok: false, error: "Geçersiz form." }, { status: 400 }); }
  const dosya = form.get("dosya");
  if (!(dosya instanceof File)) return NextResponse.json({ ok: false, error: "Dosya seçilmedi." }, { status: 400 });
  if (dosya.size <= 0 || dosya.size > EK_AZAMI_BAYT) return NextResponse.json({ ok: false, error: "Dosya boş ya da 15 MB'tan büyük." }, { status: 400 });
  const foy = form.get("foy") === "1";
  try {
    const is = await isBul(params.id);
    if (!is) return NextResponse.json({ ok: false, error: "İş bulunamadı." }, { status: 404 });
    const mime = dosya.type || "application/octet-stream";
    if (foy && mime !== "application/pdf") return NextResponse.json({ ok: false, error: "Föy olarak yalnızca PDF yüklenebilir." }, { status: 400 });
    const veri = new Uint8Array(await dosya.arrayBuffer());
    const yol = await ekKaydet(is.id, dosya.name, mime, veri);
    if (!yol) return NextResponse.json({ ok: false, error: "Dosya kaydedilemedi." }, { status: 502 });
    const ek: IsEk = {
      id: Math.random().toString(36).slice(2, 10), ad: dosya.name.slice(0, 120), yol, tur: ekTuru(mime), boyut: dosya.size,
      at: new Date().toISOString(), by: y.user.name,
    };
    const ekler = [...is.ekler, ek];
    const guncel = await isGuncelle(is.id, { ekler, ...(foy ? { foyYol: yol } : {}) }, y.user.name, true);
    await olayEkle(is.id, foy ? "foy" : "ek", foy ? `Föy yüklendi: ${ek.ad}` : `Ek yüklendi: ${ek.ad}`, y.user.name, { ekId: ek.id });
    return NextResponse.json({ ok: true, ek, is: guncel });
  } catch (e) {
    console.error("Ek yüklenemedi:", e);
    return hataYaniti(e);
  }
}

// DELETE ?ek=<id>: eki kaldırır (föy ise föy bağlantısı da düşer).
export async function DELETE(req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const ekId = req.nextUrl.searchParams.get("ek") || "";
  try {
    const is = await isBul(params.id);
    if (!is) return NextResponse.json({ ok: false, error: "İş bulunamadı." }, { status: 404 });
    const ek = is.ekler.find((e) => e.id === ekId);
    if (!ek) return NextResponse.json({ ok: false, error: "Ek bulunamadı." }, { status: 404 });
    await ekSil(ek.yol);
    const guncel = await isGuncelle(is.id, { ekler: is.ekler.filter((e) => e.id !== ekId), ...(is.foyYol === ek.yol ? { foyYol: null } : {}) }, y.user.name, true);
    await olayEkle(is.id, "ek", `Ek silindi: ${ek.ad}`, y.user.name);
    return NextResponse.json({ ok: true, is: guncel });
  } catch (e) {
    return hataYaniti(e);
  }
}
