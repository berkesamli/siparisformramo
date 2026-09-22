import { NextRequest, NextResponse } from "next/server";
import { ekOku, ekYoluGecerli } from "@/lib/mesaj/ek";
import { mesajKullanici } from "@/lib/mesaj/yetki";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Özel blob'daki mesaj ekini oturumlu ve yetkili kullanıcıya akıtır.
export async function GET(req: NextRequest) {
  const y = await mesajKullanici();
  if ("hata" in y) return y.hata;
  const p = req.nextUrl.searchParams.get("p") || "";
  if (!ekYoluGecerli(p)) return NextResponse.json({ ok: false, error: "Geçersiz ek yolu." }, { status: 400 });
  const r = await ekOku(p);
  if (!r) return NextResponse.json({ ok: false, error: "Ek bulunamadı." }, { status: 404 });
  const ad = p.split("/").pop() || "ek";
  const indir = req.nextUrl.searchParams.get("indir") === "1";
  return new NextResponse(r.stream, {
    status: 200,
    headers: {
      "Content-Type": r.contentType,
      "Content-Disposition": `${indir ? "attachment" : "inline"}; filename*=UTF-8''${encodeURIComponent(ad)}`,
      "Cache-Control": "private, max-age=3600",
      ...(r.size ? { "Content-Length": String(r.size) } : {}),
    },
  });
}
