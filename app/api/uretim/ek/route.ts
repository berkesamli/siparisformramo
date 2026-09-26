import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici } from "@/lib/uretim/yetki";
import { ekOku, ekYoluGecerli } from "@/lib/uretim/ek";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Özel blob'daki üretim ekini (görsel, PDF, föy) oturumlu ve yetkili kullanıcıya akıtır.
const INLINE_MIME = /^(image\/(jpeg|png|gif|webp)|application\/pdf)$/i;
export async function GET(req: NextRequest) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const p = req.nextUrl.searchParams.get("p") || "";
  if (!ekYoluGecerli(p)) return NextResponse.json({ ok: false, error: "Geçersiz ek yolu." }, { status: 400 });
  const r = await ekOku(p);
  if (!r) return NextResponse.json({ ok: false, error: "Ek bulunamadı." }, { status: 404 });
  const ad = p.split("/").pop() || "ek";
  const mime = (r.contentType || "").split(";")[0].trim().toLowerCase();
  const guvenli = INLINE_MIME.test(mime);
  const inline = guvenli && req.nextUrl.searchParams.get("indir") !== "1";
  return new NextResponse(r.stream, {
    status: 200,
    headers: {
      "Content-Type": guvenli ? mime : "application/octet-stream",
      "Content-Disposition": `${inline ? "inline" : "attachment"}; filename*=UTF-8''${encodeURIComponent(ad)}`,
      "X-Content-Type-Options": "nosniff",
      "Content-Security-Policy": "default-src 'none'; sandbox",
      "Cache-Control": "private, max-age=600",
      ...(r.size ? { "Content-Length": String(r.size) } : {}),
    },
  });
}
