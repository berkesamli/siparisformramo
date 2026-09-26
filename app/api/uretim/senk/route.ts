import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { perakendeSenk, perakendeTamSenk } from "@/lib/uretim/perakende-senk";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 120;

// Elle senkron: ?kaynak=perakende (varsayılan) | online ; ?tam=1 → tam tarama.
export async function POST(req: NextRequest) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const kaynak = req.nextUrl.searchParams.get("kaynak") || "perakende";
  const tam = req.nextUrl.searchParams.get("tam") === "1";
  try {
    if (kaynak === "online") {
      const { ikasSenk } = await import("@/lib/ikas/senk");
      return NextResponse.json({ ok: true, kaynak, sonuc: await ikasSenk(y.user.name, { tam }) });
    }
    const sonuc = tam ? await perakendeTamSenk(y.user.name) : await perakendeSenk(y.user.name, 3, true);
    return NextResponse.json({ ok: true, kaynak: "perakende", sonuc });
  } catch (e) {
    console.error("Senkron hatası:", e);
    return hataYaniti(e);
  }
}
