import { NextRequest, NextResponse } from "next/server";
import { ekOku, ekYoluGecerli } from "@/lib/mesaj/ek";
import { mesajKullanici } from "@/lib/mesaj/yetki";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Özel blob'daki mesaj ekini oturumlu ve yetkili kullanıcıya akıtır.
const INLINE_MIME = /^(image\/(jpeg|png|gif|webp)|application\/pdf|audio\/(ogg|mpeg|mp4|aac|wav)|video\/mp4)$/i;
export async function GET(req: NextRequest) {
  const y = await mesajKullanici();
  if ("hata" in y) return y.hata;
  const p = req.nextUrl.searchParams.get("p") || "";
  if (!ekYoluGecerli(p)) return NextResponse.json({ ok: false, error: "Geçersiz ek yolu." }, { status: 400 });
  const r = await ekOku(p);
  if (!r) return NextResponse.json({ ok: false, error: "Ek bulunamadı." }, { status: 404 });
  const ad = p.split("/").pop() || "ek";
  // Tarayıcının betik çalıştıramayacağı türler tarayıcıda açılır; gerisi (html, svg, …) indirme + octet-stream.
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
      "Cache-Control": "private, max-age=3600",
      ...(r.size ? { "Content-Length": String(r.size) } : {}),
    },
  });
}
