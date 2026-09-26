import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { isBul } from "@/lib/uretim/db";
import { ekOku } from "@/lib/uretim/ek";
import { foyUret, foyUretVeKaydet } from "@/lib/uretim/foy";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

const TR_ASCII: Record<string, string> = { ç: "c", Ç: "C", ğ: "g", Ğ: "G", ı: "i", İ: "I", ö: "o", Ö: "O", ş: "s", Ş: "S", ü: "u", Ü: "U" };
const asciiAd = (s: string) => s.replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR_ASCII[c] || c).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Za-z0-9 _-]/g, "").trim().replace(/\s+/g, "_");

function pdfYaniti(body: BodyInit, ad: string, inline: boolean, size?: number) {
  return new NextResponse(body, {
    status: 200,
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `${inline ? "inline" : "attachment"}; filename="${asciiAd(ad) || "foy"}.pdf"; filename*=UTF-8''${encodeURIComponent(ad + ".pdf")}`,
      "X-Content-Type-Options": "nosniff",
      "Cache-Control": "private, no-store",
      ...(size ? { "Content-Length": String(size) } : {}),
    },
  });
}

// GET: föyü akıtır. Kayıtlı föy (üretilmiş ya da yüklenmiş) varsa onu, yoksa anında üretir
// (üretileni Blob'a da yazar). ?indir=1 → dosya olarak indir; ?yenile=1 → kayıtlı föyü yok say.
export async function GET(req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const indir = req.nextUrl.searchParams.get("indir") === "1";
  const yenile = req.nextUrl.searchParams.get("yenile") === "1";
  try {
    const is = await isBul(params.id);
    if (!is) return NextResponse.json({ ok: false, error: "İş bulunamadı." }, { status: 404 });
    const ad = `olga_foy_${is.musteriAd || is.baslik || is.id}`;
    if (is.foyYol && !yenile) {
      const r = await ekOku(is.foyYol);
      if (r) return pdfYaniti(r.stream, ad, !indir, r.size);
    }
    const r = await foyUretVeKaydet(is, y.user.name);
    if (!r) {
      const pdf = await foyUret(is);
      if (!pdf) return NextResponse.json({ ok: false, error: "Bu işin kalemlerinde föy üretecek ölçü verisi yok. Föyü yükleyebilirsiniz." }, { status: 422 });
      return pdfYaniti(new Uint8Array(pdf), ad, !indir, pdf.length);
    }
    return pdfYaniti(new Uint8Array(r.pdf), ad, !indir, r.pdf.length);
  } catch (e) {
    console.error("Föy üretilemedi:", e);
    return hataYaniti(e);
  }
}

// POST: föyü yeniden üretir ve kaydeder (yüklenmiş föyü ezer).
export async function POST(_req: NextRequest, { params }: { params: { id: string } }) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  try {
    const is = await isBul(params.id);
    if (!is) return NextResponse.json({ ok: false, error: "İş bulunamadı." }, { status: 404 });
    const r = await foyUretVeKaydet(is, y.user.name);
    if (!r) return NextResponse.json({ ok: false, error: "Bu işin kalemlerinde föy üretecek ölçü verisi yok." }, { status: 422 });
    return NextResponse.json({ ok: true, foyYol: r.yol, is: await isBul(is.id) });
  } catch (e) {
    console.error("Föy yeniden üretilemedi:", e);
    return hataYaniti(e);
  }
}
