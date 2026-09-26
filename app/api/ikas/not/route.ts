import { NextRequest, NextResponse } from "next/server";
import { timingSafeEqual } from "node:crypto";
import { ayarYaz, dbConfigured, isBulKaynak } from "@/lib/uretim/db";
import { ikasWebhookKey } from "@/lib/ikas/client";
import { notBloklari } from "@/lib/ikas/not";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Çerçeve hesaplayıcının (olgacerceve.com sunucusu) sipariş detay metnini bize
// doğrudan bırakması için uç. ikas zaman çizelgesine yazılan not API'den
// okunamadığı için aynı metin buraya da POST edilir:
//   POST /api/ikas/not?k=<IKAS_WEBHOOK_KEY>  { "orderNumber": "1463", "metin": "CERCEVE SIPARIS DETAYI (OZL-…) …" }
// Metin OZL kimliğine (ve sipariş numarasına) göre saklanır; sipariş senkronda
// çözülürken kalemlere katılır. Sipariş zaten takvimdeyse yeniden çözülür.
export async function POST(req: NextRequest) {
  const beklenen = ikasWebhookKey();
  const gelen = req.nextUrl.searchParams.get("k") || req.headers.get("x-webhook-key") || "";
  const a = Buffer.from(gelen), b = Buffer.from(beklenen);
  if (!beklenen || a.length !== b.length || !timingSafeEqual(a, b)) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  if (!dbConfigured()) return NextResponse.json({ ok: false, error: "DATABASE_URL yok." }, { status: 503 });
  const g = (await req.json().catch(() => null)) as { orderNumber?: string; orderId?: string; ozl?: string; metin?: string } | null;
  const metin = String(g?.metin || "").slice(0, 20000).trim();
  if (!metin) return NextResponse.json({ ok: false, error: "metin gerekli." }, { status: 400 });
  const bloklar = notBloklari(metin);
  const anahtarlar: string[] = [];
  for (const bl of bloklar) if (bl.ozl) { await ayarYaz(`ikasnot:${bl.ozl}`, bl.metin); anahtarlar.push(bl.ozl); }
  if (g?.ozl) { await ayarYaz(`ikasnot:${String(g.ozl).toUpperCase()}`, metin); anahtarlar.push(String(g.ozl).toUpperCase()); }
  const no = String(g?.orderNumber || "").replace(/^#/, "").trim();
  if (no) { await ayarYaz(`ikasnot:no:${no}`, metin); anahtarlar.push(`no:${no}`); }
  // Sipariş takvimde varsa ve ikas bağlıysa yeniden çözümle (kalemler tamamlansın)
  let yenidenCozuldu = false;
  if (no) {
    try {
      const mevcut = await isBulKaynak("online", no);
      const { ikasConfigured } = await import("@/lib/ikas/client");
      if (mevcut && ikasConfigured()) {
        const { ikasSiparisNoIsle } = await import("@/lib/ikas/senk");
        const r = await ikasSiparisNoIsle(no, "hesaplayici");
        yenidenCozuldu = Boolean(r);
      }
    } catch (e) {
      console.warn("Not sonrası yeniden çözüm yapılamadı:", (e as Error)?.message);
    }
  }
  return NextResponse.json({ ok: true, anahtarlar, yenidenCozuldu });
}
